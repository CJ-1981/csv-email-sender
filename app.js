/**
 * ============================================
 * CSV Email Sender - Client-Side Application
 * ============================================
 *
 * A client-side bulk email sender for Gmail and Outlook using CSV input.
 * No backend required - all processing happens in the browser.
 *
 * Architecture:
 * - State Management: Centralized state object
 * - CSV Parsing: Papa Parse library
 * - OAuth: Google API Client (Gmail), MSAL.js (Outlook)
 * - Email Sending: Gmail API, Microsoft Graph API
 *
 * Functional Requirements Coverage:
 * - F001: CSV File Upload Processing
 * - F002: Attachment File Upload
 * - F003: Email Provider Selection
 * - F004: OAuth Authentication
 * - F005: Delay Configuration
 * - F006: Email Sending Process
 * - F007: Process Abortion
 * - F008: Completion Summary
 * - F009: Anti-Spam Warnings
 * - F010: CSV Column Validation
 * - F011: Gmail API Integration
 * - F012: Outlook API Integration
 * - F013: Progress Tracking
 * - F014: Error Handling
 *
 * Non-Functional Requirements:
 * - N001: Performance (parsing < 2s, responsive UI)
 * - N002: Usability (ARIA labels, keyboard nav, WCAG AA)
 * - N003: Security (no credential storage, HTTPS, input validation)
 * - N004: Compatibility (Chrome 90+, Firefox 88+, Safari 14+, Edge 90+)
 * - N005: Reliability (error handling, validation)
 */

// ============================================
// CONFIGURATION
// ============================================

/**
 * OAuth Configuration with localStorage support
 *
 * OAuth credentials are loaded from localStorage or use placeholder values.
 * Use the setup wizard to configure your OAuth credentials.
 *
 * For setup instructions, see the OAuth Setup Wizard modal on first load.
 */
const OAUTH_CONFIG = {
    get gmail() {
        return {
            clientId: localStorage.getItem('oauth_gmail_client_id') || 'YOUR_GMAIL_CLIENT_ID_HERE',
            scopes: 'https://www.googleapis.com/auth/gmail.send',
            discoveryDocs: ['https://www.googleapis.com/discovery/v1/apis/gmail/v1/rest']
        };
    },
    get outlook() {
        return {
            clientId: localStorage.getItem('oauth_outlook_client_id') || 'YOUR_OUTLOOK_CLIENT_ID_HERE',
            scopes: ['Mail.Send'],
            authority: 'https://login.microsoftonline.com/common'
        };
    },
    isConfigured(provider) {
        if (provider === 'gmail') {
            const clientId = localStorage.getItem('oauth_gmail_client_id');
            return clientId && clientId !== 'YOUR_GMAIL_CLIENT_ID_HERE';
        } else if (provider === 'outlook') {
            const clientId = localStorage.getItem('oauth_outlook_client_id');
            return clientId && clientId !== 'YOUR_OUTLOOK_CLIENT_ID_HERE';
        }
        return false;
    },
    isAnyConfigured() {
        return this.isConfigured('gmail') || this.isConfigured('outlook');
    }
};

/**
 * Application Constants
 */
const CONSTANTS = {
    // Validation
    MIN_DELAY: 1000, // Minimum recommended delay in ms
    MAX_ATTACHMENT_SIZE: 20 * 1024 * 1024, // 20MB in bytes
    MAX_CSV_ROWS: 10000, // Maximum recommended CSV rows

    // Email validation regex (RFC 5322 compliant)
    EMAIL_REGEX: /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$/,

    // Required CSV columns (case-insensitive)
    REQUIRED_COLUMNS: ['recipient_email', 'subject'],
    OPTIONAL_COLUMNS: ['attachment_filename', 'body_content', 'attachment', 'body', 'message']
};

// ============================================
// STATE MANAGEMENT
// ============================================

/**
 * Application State
 *
 * Centralized state object to manage application data.
 * All state mutations should go through updateState() function.
 *
 * Security Note: No credentials are stored in state or localStorage.
 * OAuth tokens are kept in memory only and cleared on page refresh.
 */
const state = {
    // File data
    csvData: null, // Parsed CSV data
    csvRows: 0, // Number of CSV rows
    attachmentFiles: [], // Uploaded attachment files

    // Authentication
    selectedProvider: null, // 'gmail' or 'outlook'
    isAuthenticated: false,
    authToken: null, // OAuth access token (memory only)
    userEmail: null, // Authenticated user email

    // Configuration
    delayMs: 5000,
    randomizationEnabled: false,
    defaultBody: '',

    // Sending state
    isSending: false,
    isAborted: false,
    currentRow: 0,
    sentCount: 0,
    failedCount: 0,
    startTime: null,
    results: [] // Sending results
};

// ============================================
// DOM ELEMENTS
// ============================================

/**
 * Cached DOM element references
 *
 * Caching DOM elements improves performance by avoiding repeated
 * document.getElementById() calls.
 */
const elements = {};

/**
 * Initialize DOM element cache
 */
function initElements() {
    elements.csvFileInput = document.getElementById('csv-file-input');
    elements.attachmentFileInput = document.getElementById('attachment-file-input');
    elements.csvStatus = document.getElementById('csv-status');
    elements.attachmentStatus = document.getElementById('attachment-status');
    elements.attachmentList = document.getElementById('attachment-list');

    elements.providerSelect = document.getElementById('provider-select');
    elements.authButton = document.getElementById('auth-button');
    elements.authStatus = document.getElementById('auth-status');

    elements.delayInput = document.getElementById('delay-input');
    elements.randomizationCheckbox = document.getElementById('randomization-checkbox');
    elements.defaultBody = document.getElementById('default-body');
    elements.delayStatus = document.getElementById('delay-status');
    elements.timeEstimate = document.getElementById('time-estimate');

    elements.startButton = document.getElementById('start-button');
    elements.abortButton = document.getElementById('abort-button');
    elements.progressBar = document.getElementById('progress-bar');
    elements.progressText = document.getElementById('progress-text');
    elements.progressLog = document.getElementById('progress-log');

    elements.sentCount = document.getElementById('sent-count');
    elements.failedCount = document.getElementById('failed-count');
    elements.timeElapsed = document.getElementById('time-elapsed');
    elements.timeRemaining = document.getElementById('time-remaining');

    elements.summarySection = document.getElementById('summary-section');
    elements.summaryTotal = document.getElementById('summary-total');
    elements.summarySent = document.getElementById('summary-sent');
    elements.summaryFailed = document.getElementById('summary-failed');
    elements.summaryTime = document.getElementById('summary-time');
    elements.resultsBody = document.getElementById('results-body');
    elements.resetButton = document.getElementById('reset-button');

    // Preset delay buttons
    elements.presetButtons = document.querySelectorAll('.preset-button');
}

// ============================================
// UTILITY FUNCTIONS
// ============================================

/**
 * Format bytes to human-readable size
 * @param {number} bytes - Size in bytes
 * @returns {string} Formatted size string
 */
function formatBytes(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

/**
 * Format milliseconds to time string (HH:MM:SS)
 * @param {number} ms - Time in milliseconds
 * @returns {string} Formatted time string
 */
function formatTime(ms) {
    const totalSeconds = Math.floor(ms / 1000);
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;
    return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
}

/**
 * Validate email address format
 * @param {string} email - Email address to validate
 * @returns {boolean} True if valid email format
 */
function isValidEmail(email) {
    return CONSTANTS.EMAIL_REGEX.test(email);
}

/**
 * Calculate randomized delay based on configuration
 * @param {number} baseDelay - Base delay in milliseconds
 * @param {number} randomizationPercent - Randomization percentage (e.g., 20 for ±20%)
 * @returns {number} Randomized delay in milliseconds
 */
function calculateRandomizedDelay(baseDelay, randomizationPercent = 20) {
    if (!state.randomizationEnabled) return baseDelay;

    const randomFactor = (Math.random() - 0.5) * 2 * (randomizationPercent / 100);
    const randomizedDelay = baseDelay * (1 + randomFactor);
    return Math.max(0, Math.round(randomizedDelay));
}

/**
 * Sleep/delay function
 * @param {number} ms - Milliseconds to sleep
 * @returns {Promise} Promise that resolves after delay
 */
function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Sanitize HTML to prevent XSS attacks
 * @param {string} str - String to sanitize
 * @returns {string} Sanitized string
 */
function sanitizeHTML(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

// ============================================
// CSV PROCESSING (F001, F010)
// ============================================

/**
 * Handle CSV file upload
 *
 * Parses CSV using Papa Parse and validates required columns.
 * Validates presence of required columns: recipient_email, subject
 *
 * Functional Requirements: F001, F010
 * Acceptance Criteria: AC001, AC002
 */
function handleCSVUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    setStatus(elements.csvStatus, 'info', 'Parsing CSV...');

    Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: function(results) {
            processCSVResults(results);
        },
        error: function(error) {
            setStatus(elements.csvStatus, 'error', `CSV parsing error: ${error.message}`);
            logMessage('error', `Failed to parse CSV: ${error.message}`);
        }
    });
}

/**
 * Process parsed CSV results
 * @param {Object} results - Papa Parse results object
 */
function processCSVResults(results) {
    const { data, meta } = results;

    // Check for empty CSV (AC031)
    if (!data || data.length === 0) {
        setStatus(elements.csvStatus, 'error', 'CSV file is empty');
        logMessage('error', 'CSV file is empty (0 rows)');
        state.csvData = null;
        state.csvRows = 0;
        updateStartButton();
        return;
    }

    // Validate required columns (AC002)
    const columns = meta.fields.map(col => col.toLowerCase());
    const requiredCols = CONSTANTS.REQUIRED_COLUMNS.map(col => col.toLowerCase());
    const missingColumns = requiredCols.filter(col => !columns.includes(col));

    if (missingColumns.length > 0) {
        setStatus(elements.csvStatus, 'error', `Missing required columns: ${missingColumns.join(', ')}`);
        logMessage('error', `CSV validation failed: Missing columns ${missingColumns.join(', ')}`);
        state.csvData = null;
        state.csvRows = 0;
        updateStartButton();
        return;
    }

    // Validate email addresses (AC035)
    const invalidEmails = [];
    data.forEach((row, index) => {
        const email = getRecipientEmail(row);
        if (!isValidEmail(email)) {
            invalidEmails.push({ row: index + 1, email, actual: email });
        }
    });

    if (invalidEmails.length > 0) {
        const emailList = invalidEmails.slice(0, 5).map(e => `Row ${e.row}: '${e.actual}'`).join('\n');
        const moreText = invalidEmails.length > 5 ? `\n... and ${invalidEmails.length - 5} more` : '';
        setStatus(elements.csvStatus, 'warning', `${invalidEmails.length} invalid email addresses found`);
        logMessage('warning', `Invalid emails found:\n${emailList}${moreText}`);

        // Ask user to confirm
        if (!confirm(`${invalidEmails.length} invalid email addresses found. Continue anyway?`)) {
            state.csvData = null;
            state.csvRows = 0;
            updateStartButton();
            return;
        }
    }

    // CSV is valid
    state.csvData = data;
    state.csvRows = data.length;
    setStatus(elements.csvStatus, 'success', `CSV parsed successfully: ${data.length} rows`);
    logMessage('info', `CSV parsed: ${data.length} rows loaded`);

    // Check if CSV exceeds recommended size
    if (data.length > CONSTANTS.MAX_CSV_ROWS) {
        logMessage('warning', `CSV has ${data.length} rows, which exceeds recommended limit of ${CONSTANTS.MAX_CSV_ROWS}. Performance may be impacted.`);
    }

    updateTimeEstimate();
    updateStartButton();
}

/**
 * Get recipient email from CSV row
 * Handles multiple column name variations
 * @param {Object} row - CSV row object
 * @returns {string} Recipient email address
 */
function getRecipientEmail(row) {
    const variations = ['recipient_email', 'email', 'to'];
    for (const col of variations) {
        if (row[col]) return row[col].trim();
    }
    return '';
}

/**
 * Get subject from CSV row
 * @param {Object} row - CSV row object
 * @returns {string} Email subject
 */
function getSubject(row) {
    const variations = ['subject'];
    for (const col of variations) {
        if (row[col]) return row[col].trim();
    }
    return '';
}

/**
 * Get attachment filename from CSV row
 * @param {Object} row - CSV row object
 * @returns {string} Attachment filename
 */
function getAttachmentFilename(row) {
    const variations = ['attachment_filename', 'attachment'];
    for (const col of variations) {
        if (row[col]) return row[col].trim();
    }
    return '';
}

/**
 * Get body content from CSV row
 * @param {Object} row - CSV row object
 * @returns {string} Body content
 */
function getBodyContent(row) {
    const variations = ['body_content', 'body', 'message'];
    for (const col of variations) {
        if (row[col]) return row[col].trim();
    }
    return '';
}

// ============================================
// ATTACHMENT HANDLING (F002)
// ============================================

/**
 * Handle attachment file upload
 *
 * Stores files in browser memory as File objects.
 * Validates total attachment size (warns if > 20MB).
 *
 * Functional Requirements: F002
 * Acceptance Criteria: AC003, AC004
 */
function handleAttachmentUpload(event) {
    const files = Array.from(event.target.files);
    if (files.length === 0) return;

    state.attachmentFiles = files;

    // Calculate total size
    const totalSize = files.reduce((sum, file) => sum + file.size, 0);

    // Display attachment list
    displayAttachmentList(files);

    // Warn if exceeds recommended size (AC004)
    if (totalSize > CONSTANTS.MAX_ATTACHMENT_SIZE) {
        setStatus(elements.attachmentStatus, 'warning',
            `Total attachment size (${formatBytes(totalSize)}) exceeds recommended limit (${formatBytes(CONSTANTS.MAX_ATTACHMENT_SIZE)}). Consider reducing sizes.`);
        logMessage('warning', `Attachment size ${formatBytes(totalSize)} exceeds recommended limit`);
    } else {
        setStatus(elements.attachmentStatus, 'success',
            `${files.length} file(s) selected, total size: ${formatBytes(totalSize)}`);
    }
}

/**
 * Display attachment list in UI
 * @param {File[]} files - Array of File objects
 */
function displayAttachmentList(files) {
    elements.attachmentList.innerHTML = '';

    files.forEach(file => {
        const item = document.createElement('div');
        item.className = 'attachment-item';
        item.innerHTML = `
            <span class="attachment-name">${sanitizeHTML(file.name)}</span>
            <span class="attachment-size">${formatBytes(file.size)}</span>
        `;
        elements.attachmentList.appendChild(item);
    });
}

/**
 * Find attachment file by filename
 * @param {string} filename - Attachment filename
 * @returns {File|null} File object or null if not found
 */
function findAttachment(filename) {
    if (!filename) return null;
    return state.attachmentFiles.find(file => file.name === filename) || null;
}

// ============================================
// PROVIDER SELECTION (F003)
// ============================================

/**
 * Handle provider selection change
 *
 * Functional Requirements: F003
 * Acceptance Criteria: AC005
 */
function handleProviderChange() {
    const provider = elements.providerSelect.value;
    state.selectedProvider = provider || null;
    state.isAuthenticated = false;
    state.authToken = null;
    state.userEmail = null;

    // Update auth button
    if (provider === 'gmail') {
        elements.authButton.textContent = 'Authenticate with Gmail';
        elements.authButton.disabled = false;
        setStatus(elements.authStatus, 'info', 'Click to authenticate with Gmail');
    } else if (provider === 'outlook') {
        elements.authButton.textContent = 'Authenticate with Outlook';
        elements.authButton.disabled = false;
        setStatus(elements.authStatus, 'info', 'Click to authenticate with Outlook');
    } else {
        elements.authButton.textContent = 'Authenticate';
        elements.authButton.disabled = true;
        setStatus(elements.authStatus, '', '');
    }

    updateStartButton();
}

// ============================================
// OAUTH AUTHENTICATION (F004)
// ============================================

/**
 * Initialize OAuth authentication
 *
 * Functional Requirements: F004
 * Acceptance Criteria: AC006, AC007
 */
async function authenticate() {
    const provider = state.selectedProvider;
    if (!provider) return;

    logMessage('info', `Initiating ${provider} authentication...`);

    try {
        if (provider === 'gmail') {
            await authenticateGmail();
        } else if (provider === 'outlook') {
            await authenticateOutlook();
        }
    } catch (error) {
        setStatus(elements.authStatus, 'error', `Authentication failed: ${error.message}`);
        logMessage('error', `Authentication error: ${error.message}`);
    }
}

/**
 * Authenticate with Gmail using Google API Client
 *
 * Uses OAuth 2.0 implicit flow for client-side authentication.
 * Requests gmail.send scope for sending emails.
 */
async function authenticateGmail() {
    // Get current OAuth config (reads from localStorage)
    const oauthConfig = OAUTH_CONFIG.gmail;

    // Check if Client ID is configured
    if (!oauthConfig.clientId || oauthConfig.clientId === 'YOUR_GMAIL_CLIENT_ID_HERE') {
        // Show setup wizard if not configured
        openOAuthWizard();
        throw new Error('Gmail Client ID not configured. Please complete the setup wizard.');
    }

    // Initialize Google API Client
    await gapi.load('client:auth2', async () => {
        try {
            await gapi.auth2.init({
                client_id: oauthConfig.clientId,
                scope: oauthConfig.scopes
            });

            const authInstance = gapi.auth2.getAuthInstance();
            const user = await authInstance.signIn();

            // Get access token
            const authResponse = user.getAuthResponse();
            state.authToken = authResponse.access_token;
            state.userEmail = user.getBasicProfile().getEmail();
            state.isAuthenticated = true;

            setStatus(elements.authStatus, 'success', `Authenticated as: ${state.userEmail}`);
            logMessage('success', `Gmail authentication successful: ${state.userEmail}`);
            elements.authButton.disabled = true;
            updateStartButton();

        } catch (error) {
            throw new Error(`Gmail authentication failed: ${error.error || error.message}`);
        }
    });
}

/**
 * Authenticate with Outlook using MSAL.js
 *
 * Uses OAuth 2.0 authorization code flow with PKCE.
 * Requests Mail.Send scope for sending emails.
 */
async function authenticateOutlook() {
    // Get current OAuth config (reads from localStorage)
    const oauthConfig = OAUTH_CONFIG.outlook;

    // Check if Client ID is configured
    if (!oauthConfig.clientId || oauthConfig.clientId === 'YOUR_OUTLOOK_CLIENT_ID_HERE') {
        // Show setup wizard if not configured
        openOAuthWizard();
        throw new Error('Outlook Client ID not configured. Please complete the setup wizard.');
    }

    // Initialize MSAL
    const msalConfig = {
        auth: {
            clientId: oauthConfig.clientId,
            authority: oauthConfig.authority
        }
    };

    const msalInstance = new msal.PublicClientApplication(msalConfig);
    await msalInstance.initialize();

    const loginRequest = {
        scopes: oauthConfig.scopes
    };

    try {
        const loginResponse = await msalInstance.loginPopup(loginRequest);
        state.authToken = loginResponse.accessToken;
        state.userEmail = loginResponse.account.username;
        state.isAuthenticated = true;

        setStatus(elements.authStatus, 'success', `Authenticated as: ${state.userEmail}`);
        logMessage('success', `Outlook authentication successful: ${state.userEmail}`);
        elements.authButton.disabled = true;
        updateStartButton();

    } catch (error) {
        throw new Error(`Outlook authentication failed: ${error.message}`);
    }
}

// ============================================
// DELAY CONFIGURATION (F005)
// ============================================

/**
 * Handle delay input change
 *
 * Functional Requirements: F005
 * Acceptance Criteria: AC008, AC009
 */
function handleDelayChange() {
    const delay = parseInt(elements.delayInput.value) || 0;
    state.delayMs = delay;

    // Show warning if delay is too low (AC009)
    if (delay < CONSTANTS.MIN_DELAY && delay > 0) {
        setStatus(elements.delayStatus, 'warning',
            `Delay less than ${CONSTANTS.MIN_DELAY}ms may trigger rate limits. Recommended minimum: ${CONSTANTS.MIN_DELAY}ms`);
    } else if (delay === 0) {
        setStatus(elements.delayStatus, 'error',
            'Zero delay may trigger rate limits and account suspension. Recommended minimum: 1000ms');
    } else {
        setStatus(elements.delayStatus, '', '');
    }

    updateTimeEstimate();
}

/**
 * Handle randomization toggle
 */
function handleRandomizationChange() {
    state.randomizationEnabled = elements.randomizationCheckbox.checked;
    updateTimeEstimate();
}

/**
 * Update time estimate display
 */
function updateTimeEstimate() {
    if (!state.csvRows || state.csvRows === 0) {
        elements.timeEstimate.innerHTML = '<strong>Estimated Total Time:</strong> Not calculated yet';
        return;
    }

    const avgDelay = state.randomizationEnabled ? state.delayMs : state.delayMs;
    const totalTimeMs = avgDelay * state.csvRows;
    const timeStr = formatTime(totalTimeMs);

    elements.timeEstimate.innerHTML = `<strong>Estimated Total Time:</strong> ${timeStr} (${state.csvRows} emails × ${avgDelay}ms delay)`;
}

// ============================================
// SENDING PROCESS (F006, F011, F012)
// ============================================

/**
 * Start sending process
 *
 * Functional Requirements: F006
 * Acceptance Criteria: AC010
 */
async function startSending() {
    // Validate preconditions (F006)
    if (!state.csvData || state.csvRows === 0) {
        alert('Please upload a valid CSV file first.');
        return;
    }

    if (!state.isAuthenticated || !state.authToken) {
        alert('Please authenticate with your email provider first.');
        return;
    }

    if (state.delayMs === 0) {
        // Warning for zero delay (AC034)
        if (!confirm('Zero delay may trigger rate limits and account suspension. Recommended minimum: 1000ms. Continue?')) {
            return;
        }
    }

    // Check for missing attachments (AC032)
    const missingAttachments = checkMissingAttachments();
    if (missingAttachments.length > 0) {
        const fileList = missingAttachments.slice(0, 5).join(', ') +
            (missingAttachments.length > 5 ? `... and ${missingAttachments.length - 5} more` : '');
        if (!confirm(`CSV references attachments not uploaded: ${fileList}\n\nContinue anyway?`)) {
            return;
        }
    }

    // Confirm before starting
    if (!confirm(`Ready to send ${state.csvRows} emails with ${state.delayMs}ms delay between sends.\n\nProceed?`)) {
        return;
    }

    // Initialize sending state
    state.isSending = true;
    state.isAborted = false;
    state.currentRow = 0;
    state.sentCount = 0;
    state.failedCount = 0;
    state.startTime = Date.now();
    state.results = [];

    // Update UI
    elements.startButton.disabled = true;
    elements.abortButton.disabled = false;
    elements.csvFileInput.disabled = true;
    elements.attachmentFileInput.disabled = true;
    elements.providerSelect.disabled = true;
    elements.delayInput.disabled = true;
    elements.progressBar.classList.add('sending');

    logMessage('info', `Starting email sending process: ${state.csvRows} emails`);

    // Send emails
    await sendEmails();

    // Show summary
    showSummary();
}

/**
 * Check for missing attachments referenced in CSV
 * @returns {string[]} Array of missing attachment filenames
 */
function checkMissingAttachments() {
    const missing = new Set();

    state.csvData.forEach((row) => {
        const attachmentFilename = getAttachmentFilename(row);
        if (attachmentFilename && !findAttachment(attachmentFilename)) {
            missing.add(attachmentFilename);
        }
    });

    return Array.from(missing);
}

/**
 * Send emails one by one with delays
 *
 * Functional Requirements: F006, F011, F012, F013, F014
 * Acceptance Criteria: AC010, AC011, AC012, AC014, AC015
 */
async function sendEmails() {
    for (let i = 0; i < state.csvRows; i++) {
        // Check for abort
        if (state.isAborted) {
            logMessage('warning', 'Process aborted by user');
            break;
        }

        state.currentRow = i + 1;
        const row = state.csvData[i];

        // Update progress (F013)
        updateProgress();

        try {
            // Send email
            const result = await sendSingleEmail(row);

            if (result.success) {
                state.sentCount++;
                state.results.push({
                    row: i + 1,
                    recipient: result.recipient,
                    subject: result.subject,
                    status: 'success',
                    error: null
                });
                logMessage('success', `[${i + 1}/${state.csvRows}] ✓ Sent to ${result.recipient}`);
            } else {
                state.failedCount++;
                state.results.push({
                    row: i + 1,
                    recipient: result.recipient,
                    subject: result.subject,
                    status: 'failed',
                    error: result.error
                });
                logMessage('error', `[${i + 1}/${state.csvRows}] ✗ Failed: ${result.error}`);
            }

        } catch (error) {
            state.failedCount++;
            state.results.push({
                row: i + 1,
                recipient: getRecipientEmail(row),
                subject: getSubject(row),
                status: 'failed',
                error: error.message
            });
            logMessage('error', `[${i + 1}/${state.csvRows}] ✗ Error: ${error.message}`);
        }

        // Update stats
        updateStats();

        // Apply delay before next send (except for last email)
        if (i < state.csvRows - 1 && !state.isAborted) {
            const delay = calculateRandomizedDelay(state.delayMs);
            await sleep(delay);
        }
    }

    state.isSending = false;
    elements.progressBar.classList.remove('sending');
    logMessage('info', 'Sending process completed');
}

/**
 * Send a single email
 *
 * Functional Requirements: F011 (Gmail), F012 (Outlook)
 * Acceptance Criteria: AC011, AC012, AC014, AC015
 */
async function sendSingleEmail(row) {
    const recipient = getRecipientEmail(row);
    const subject = getSubject(row);
    const body = getBodyContent(row) || state.defaultBody;
    const attachmentFilename = getAttachmentFilename(row);
    const attachment = findAttachment(attachmentFilename);

    try {
        if (state.selectedProvider === 'gmail') {
            return await sendGmailEmail(recipient, subject, body, attachment);
        } else if (state.selectedProvider === 'outlook') {
            return await sendOutlookEmail(recipient, subject, body, attachment);
        }
    } catch (error) {
        // Handle specific error types (F014, AC012, AC013, AC015)
        if (error.status === 429) {
            throw new Error('Rate limit exceeded - Daily quota exhausted');
        } else if (error.status === 401) {
            throw new Error('Authentication failed - Token may have expired');
        } else if (error.status === 400) {
            throw new Error(`Invalid request: ${error.message}`);
        } else {
            throw new Error(error.message || 'Unknown error occurred');
        }
    }
}

/**
 * Send email via Gmail API
 *
 * Functional Requirements: F011
 * Acceptance Criteria: AC011, AC012, AC013
 *
 * Constructs RFC 2822 formatted email with base64-encoded attachments.
 */
async function sendGmailEmail(to, subject, body, attachment) {
    // Build RFC 2822 email
    let emailContent = [
        `To: ${to}`,
        `Subject: ${subject}`,
        'MIME-Version: 1.0',
        'Content-Type: text/plain; charset=utf-8'
    ];

    if (attachment) {
        // Email with attachment
        const boundary = 'boundary_' + Date.now();
        emailContent = [
            `To: ${to}`,
            `Subject: ${subject}`,
            'MIME-Version: 1.0',
            `Content-Type: multipart/mixed; boundary="${boundary}"`,
            '',
            `--${boundary}`,
            'Content-Type: text/plain; charset=utf-8',
            '',
            body,
            ''
        ];

        // Read attachment as base64
        const base64Attachment = await fileToBase64(attachment);
        emailContent.push(
            `--${boundary}`,
            'Content-Type: application/octet-stream',
            `Content-Disposition: attachment; filename="${attachment.name}"`,
            'Content-Transfer-Encoding: base64',
            '',
            base64Attachment,
            `--${boundary}--`
        );
    } else {
        // Plain text email
        emailContent.push('', body);
    }

    const email = emailContent.join('\r\n');
    // Encode email for Gmail API: UTF-8 bytes to base64
    // Using TextEncoder to avoid deprecated unescape()
    const encoder = new TextEncoder();
    const uint8Array = encoder.encode(email);
    const binaryString = Array.from(uint8Array, byte => String.fromCharCode(byte)).join('');
    const base64Email = btoa(binaryString);

    // Send via Gmail API
    const response = await fetch('https://gmail.googleapis.com/gmail/v1/users/me/messages/send', {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${state.authToken}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ raw: base64Email })
    });

    if (!response.ok) {
        const error = await response.json();
        throw { status: response.status, message: error.error?.message || 'Gmail API error' };
    }

    const result = await response.json();
    return { success: true, recipient: to, subject: subject, id: result.id };
}

/**
 * Send email via Microsoft Graph API (Outlook)
 *
 * Functional Requirements: F012
 * Acceptance Criteria: AC014, AC015
 *
 * Constructs MIME message with file attachments.
 */
async function sendOutlookEmail(to, subject, body, attachment) {
    const message = {
        subject: subject,
        toRecipients: [{
            emailAddress: { address: to }
        }],
        body: {
            contentType: 'Text',
            content: body
        }
    };

    // Add attachment if present
    if (attachment) {
        const base64Attachment = await fileToBase64(attachment);
        message.attachments = [{
            '@odata.type': '#microsoft.graph.fileAttachment',
            name: attachment.name,
            contentBytes: base64Attachment
        }];
    }

    // Send via Microsoft Graph API
    const response = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${state.authToken}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ message: message })
    });

    if (!response.ok) {
        const error = await response.json();
        throw { status: response.status, message: error.error?.message || 'Outlook API error' };
    }

    return { success: true, recipient: to, subject: subject };
}

/**
 * Convert file to base64 string
 * @param {File} file - File object
 * @returns {Promise<string>} Base64 encoded string
 */
function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
            const base64 = reader.result.split(',')[1];
            resolve(base64);
        };
        reader.onerror = error => reject(error);
        reader.readAsDataURL(file);
    });
}

// ============================================
// PROGRESS TRACKING (F013)
// ============================================

/**
 * Update progress display
 *
 * Functional Requirements: F013
 * Acceptance Criteria: AC010
 */
function updateProgress() {
    const percentage = Math.round((state.currentRow / state.csvRows) * 100);
    elements.progressBar.style.width = `${percentage}%`;
    elements.progressText.textContent = `${state.currentRow} / ${state.csvRows} (${percentage}%)`;
    elements.progressBar.parentElement.setAttribute('aria-valuenow', percentage);
}

/**
 * Update statistics display
 */
function updateStats() {
    elements.sentCount.textContent = state.sentCount;
    elements.failedCount.textContent = state.failedCount;

    // Time elapsed
    const elapsed = Date.now() - state.startTime;
    elements.timeElapsed.textContent = formatTime(elapsed);

    // Estimate remaining time
    const remainingRows = state.csvRows - state.currentRow;
    const avgTimePerRow = elapsed / state.currentRow;
    const remainingTime = remainingRows * avgTimePerRow;
    elements.timeRemaining.textContent = formatTime(remainingTime);
}

// ============================================
// ABORT FUNCTIONALITY (F007)
// ============================================

/**
 * Abort sending process
 *
 * Functional Requirements: F007
 * Acceptance Criteria: AC016
 */
function abortSending() {
    if (state.isSending) {
        state.isAborted = true;
        logMessage('warning', 'Aborting sending process...');
        elements.abortButton.disabled = true;
    }
}

// ============================================
// COMPLETION SUMMARY (F008)
// ============================================

/**
 * Show completion summary
 *
 * Functional Requirements: F008
 * Acceptance Criteria: AC018, AC019, AC020
 */
function showSummary() {
    // Re-enable controls
    elements.startButton.disabled = false;
    elements.abortButton.disabled = true;
    elements.csvFileInput.disabled = false;
    elements.attachmentFileInput.disabled = false;
    elements.providerSelect.disabled = false;
    elements.delayInput.disabled = false;

    // Calculate total time
    const totalTime = Date.now() - state.startTime;

    // Update summary stats
    elements.summaryTotal.textContent = state.csvRows;
    elements.summarySent.textContent = state.sentCount;
    elements.summaryFailed.textContent = state.failedCount;
    elements.summaryTime.textContent = formatTime(totalTime);

    // Populate results table
    populateResultsTable();

    // Show summary section
    elements.summarySection.classList.remove('hidden');

    // Scroll to summary
    elements.summarySection.scrollIntoView({ behavior: 'smooth' });

    // Log completion
    if (state.isAborted) {
        logMessage('warning', `Process aborted: ${state.sentCount} sent, ${state.failedCount} failed, ${state.csvRows - state.currentRow} not attempted`);
    } else {
        logMessage('success', `Process completed: ${state.sentCount} sent, ${state.failedCount} failed`);
    }
}

/**
 * Populate results table with sending results
 */
function populateResultsTable() {
    elements.resultsBody.innerHTML = '';

    state.results.forEach(result => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${result.row}</td>
            <td>${sanitizeHTML(result.recipient)}</td>
            <td>${sanitizeHTML(result.subject)}</td>
            <td class="${result.status === 'success' ? 'status-success' : 'status-failed'}">
                ${result.status === 'success' ? '✓ Sent' : '✗ Failed'}
            </td>
            <td>${result.error ? sanitizeHTML(result.error) : ''}</td>
        `;
        elements.resultsBody.appendChild(row);
    });
}

/**
 * Reset application to start over
 */
function resetApplication() {
    // Clear state
    state.csvData = null;
    state.csvRows = 0;
    state.attachmentFiles = [];
    state.isSending = false;
    state.isAborted = false;
    state.currentRow = 0;
    state.sentCount = 0;
    state.failedCount = 0;
    state.startTime = null;
    state.results = [];

    // Reset form elements
    elements.csvFileInput.value = '';
    elements.attachmentFileInput.value = '';
    elements.providerSelect.value = '';
    elements.delayInput.value = '5000';
    elements.randomizationCheckbox.checked = false;
    elements.defaultBody.value = '';

    // Reset displays
    elements.csvStatus.textContent = '';
    elements.attachmentStatus.textContent = '';
    elements.attachmentList.innerHTML = '';
    elements.authStatus.textContent = '';
    elements.delayStatus.textContent = '';
    elements.timeEstimate.innerHTML = '<strong>Estimated Total Time:</strong> Not calculated yet';

    elements.progressBar.style.width = '0%';
    elements.progressText.textContent = '0 / 0 (0%)';
    elements.sentCount.textContent = '0';
    elements.failedCount.textContent = '0';
    elements.timeElapsed.textContent = '00:00:00';
    elements.timeRemaining.textContent = '--:--:--';

    elements.progressLog.innerHTML = '<div class="log-entry log-info">Ready to start. Upload CSV and authenticate to begin.</div>';

    // Hide summary
    elements.summarySection.classList.add('hidden');

    // Update buttons
    handleProviderChange();
}

// ============================================
// UI HELPER FUNCTIONS
// ============================================

/**
 * Set status message with type
 * @param {HTMLElement} element - Status element
 * @param {string} type - Message type (success, error, warning, info)
 * @param {string} message - Message text
 */
function setStatus(element, type, message) {
    element.textContent = message;
    element.className = 'status-message';
    if (type) {
        element.classList.add(`status-${type}`);
    }
}

/**
 * Add message to progress log
 * @param {string} type - Message type (info, success, error, warning)
 * @param {string} message - Message text
 */
function logMessage(type, message) {
    const entry = document.createElement('div');
    entry.className = `log-entry log-${type}`;
    entry.textContent = `[${new Date().toLocaleTimeString()}] ${message}`;
    elements.progressLog.appendChild(entry);
    elements.progressLog.scrollTop = elements.progressLog.scrollHeight;
}

/**
 * Update start button state based on prerequisites
 */
function updateStartButton() {
    const hasCSV = state.csvData && state.csvRows > 0;
    const isAuthenticated = state.isAuthenticated;
    elements.startButton.disabled = !(hasCSV && isAuthenticated);
}

// ============================================
// EVENT LISTENERS
// ============================================

/**
 * Initialize event listeners
 */
function initEventListeners() {
    // CSV upload
    elements.csvFileInput.addEventListener('change', handleCSVUpload);

    // Attachment upload
    elements.attachmentFileInput.addEventListener('change', handleAttachmentUpload);

    // Provider selection
    elements.providerSelect.addEventListener('change', handleProviderChange);

    // Authentication
    elements.authButton.addEventListener('click', authenticate);

    // Delay configuration
    elements.delayInput.addEventListener('input', handleDelayChange);
    elements.randomizationCheckbox.addEventListener('change', handleRandomizationChange);
    elements.defaultBody.addEventListener('input', () => {
        state.defaultBody = elements.defaultBody.value;
    });

    // Preset delay buttons
    elements.presetButtons.forEach(button => {
        button.addEventListener('click', () => {
            const delay = parseInt(button.dataset.delay);
            elements.delayInput.value = delay;
            handleDelayChange();
        });
    });

    // Sending controls
    elements.startButton.addEventListener('click', startSending);
    elements.abortButton.addEventListener('click', abortSending);

    // Reset button
    elements.resetButton.addEventListener('click', resetApplication);
}

// ============================================
// OAUTH SETUP WIZARD
// ============================================

/**
 * OAuth Setup Wizard Functions
 *
 * Manages the OAuth setup wizard modal that guides users through
 * configuring Gmail and Outlook OAuth credentials.
 */

let currentWizardStep = 'welcome';
let configuredProviders = [];

/**
 * Check if wizard should be shown on first load
 */
function shouldShowOAuthWizard() {
    // Show wizard if no OAuth credentials are configured
    return !OAUTH_CONFIG.isAnyConfigured();
}

/**
 * Show OAuth setup wizard modal
 */
function openOAuthWizard() {
    const modal = document.getElementById('oauth-wizard-modal');
    if (modal) {
        modal.classList.remove('hidden');
        modal.setAttribute('aria-hidden', 'false');

        // Reset wizard state
        currentWizardStep = 'welcome';
        configuredProviders = [];

        // Show welcome step
        showStep('welcome');

        // Populate URLs
        const currentUrl = window.location.href.split('?')[0]; // Remove query params
        const gmailOriginUrl = document.getElementById('gmail-origin-url');
        const outlookRedirectUrl = document.getElementById('outlook-redirect-url');

        if (gmailOriginUrl) gmailOriginUrl.textContent = currentUrl;
        if (outlookRedirectUrl) outlookRedirectUrl.textContent = currentUrl;

        // Prevent background scrolling
        document.body.style.overflow = 'hidden';
    }
}

/**
 * Close OAuth setup wizard modal
 */
function closeOAuthWizard() {
    const modal = document.getElementById('oauth-wizard-modal');
    if (modal) {
        modal.classList.add('hidden');
        modal.setAttribute('aria-hidden', 'true');

        // Restore background scrolling
        document.body.style.overflow = '';

        // Check if we should remind user later
        if (!OAUTH_CONFIG.isAnyConfigured()) {
            logMessage('warning', 'OAuth setup skipped. Configure credentials to use email sending features.');
        }
    }
}

/**
 * Skip wizard and remind later
 */
function skipWizard() {
    closeOAuthWizard();
    // Store flag to remind user later
    localStorage.setItem('oauth_wizard_dismissed', Date.now().toString());
}

/**
 * Show specific wizard step
 * @param {string} stepName - Step identifier
 */
function showStep(stepName) {
    // Hide all steps
    const steps = document.querySelectorAll('.wizard-step');
    steps.forEach(step => step.classList.add('hidden'));

    // Show target step
    const targetStep = document.querySelector(`[data-step="${stepName}"]`);
    if (targetStep) {
        targetStep.classList.remove('hidden');
        currentWizardStep = stepName;
    }
}

/**
 * Show provider selection step
 */
function showProviderSelection() {
    showStep('provider');
}

/**
 * Show setup instructions for specific provider
 * @param {string} provider - Provider name (gmail or outlook)
 */
function showSetupInstructions(provider) {
    if (provider === 'gmail') {
        showStep('gmail-setup');
    } else if (provider === 'outlook') {
        showStep('outlook-setup');
    }
}

/**
 * Save Gmail credentials to localStorage
 */
function saveGmailCredentials() {
    const clientIdInput = document.getElementById('gmail-client-id');
    const clientId = clientIdInput.value.trim();

    // Validate Gmail Client ID format
    if (!clientId) {
        alert('Please enter a Gmail Client ID');
        return;
    }

    // Basic format validation (should end with .apps.googleusercontent.com)
    if (!clientId.includes('.apps.googleusercontent.com')) {
        if (!confirm('The Client ID format may be incorrect. Gmail Client IDs typically end with ".apps.googleusercontent.com". Continue anyway?')) {
            return;
        }
    }

    // Save to localStorage
    localStorage.setItem('oauth_gmail_client_id', clientId);
    configuredProviders.push('gmail');

    // Show success
    showSuccessStep();
}

/**
 * Save Outlook credentials to localStorage
 */
function saveOutlookCredentials() {
    const clientIdInput = document.getElementById('outlook-client-id');
    const clientId = clientIdInput.value.trim();

    // Validate Outlook Client ID format
    if (!clientId) {
        alert('Please enter an Outlook Client ID');
        return;
    }

    // Basic format validation (should be a UUID)
    const uuidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (!uuidPattern.test(clientId)) {
        if (!confirm('The Client ID format may be incorrect. Outlook Client IDs should be a UUID. Continue anyway?')) {
            return;
        }
    }

    // Save to localStorage
    localStorage.setItem('oauth_outlook_client_id', clientId);
    configuredProviders.push('outlook');

    // Show success
    showSuccessStep();
}

/**
 * Show success step with configured providers
 */
function showSuccessStep() {
    // Update success step with configured providers
    const gmailStatus = document.getElementById('setup-gmail-status');
    const outlookStatus = document.getElementById('setup-outlook-status');

    if (gmailStatus) {
        if (configuredProviders.includes('gmail')) {
            gmailStatus.classList.remove('hidden');
        } else {
            gmailStatus.classList.add('hidden');
        }
    }

    if (outlookStatus) {
        if (configuredProviders.includes('outlook')) {
            outlookStatus.classList.remove('hidden');
        } else {
            outlookStatus.classList.add('hidden');
        }
    }

    showStep('success');
}

/**
 * Copy text to clipboard
 * @param {string} elementId - ID of element containing text to copy
 */
function copyToClipboard(elementId) {
    const element = document.getElementById(elementId);
    const text = element.textContent || element.innerText;

    navigator.clipboard.writeText(text).then(() => {
        // Visual feedback
        const button = element.parentElement.querySelector('.copy-button');
        const originalText = button.textContent;
        button.textContent = 'Copied!';
        button.style.backgroundColor = 'var(--success-color)';

        setTimeout(() => {
            button.textContent = originalText;
            button.style.backgroundColor = '';
        }, 2000);
    }).catch(err => {
        console.error('Failed to copy to clipboard:', err);
        alert('Failed to copy. Please select and copy manually.');
    });
}

/**
 * Reset OAuth configuration (for testing/reconfiguration)
 */
function resetOAuthConfiguration() {
    if (confirm('This will clear all saved OAuth credentials. Continue?')) {
        localStorage.removeItem('oauth_gmail_client_id');
        localStorage.removeItem('oauth_outlook_client_id');
        localStorage.removeItem('oauth_wizard_dismissed');
        alert('OAuth credentials cleared. Reloading page...');
        location.reload();
    }
}

// ============================================
// INITIALIZATION
// ============================================

/**
 * Initialize application
 */
function init() {
    initElements();
    initEventListeners();

    // Check if OAuth wizard should be shown
    if (shouldShowOAuthWizard()) {
        // Delay showing wizard to allow UI to fully load
        setTimeout(() => {
            openOAuthWizard();
        }, 500);
    } else {
        logMessage('info', 'OAuth credentials configured. Ready to use.');
    }

    logMessage('info', 'Application loaded. Ready to use.');
    console.log('CSV Email Sender initialized');

    // Log OAuth configuration status
    if (OAUTH_CONFIG.isConfigured('gmail')) {
        console.log('✓ Gmail OAuth configured');
    } else {
        console.log('✗ Gmail OAuth not configured');
    }

    if (OAUTH_CONFIG.isConfigured('outlook')) {
        console.log('✓ Outlook OAuth configured');
    } else {
        console.log('✗ Outlook OAuth not configured');
    }

    // Add reset function to global scope for console access
    window.resetOAuthConfiguration = resetOAuthConfiguration;
}

// Initialize when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}
