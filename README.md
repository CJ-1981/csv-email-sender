# CSV Email Sender

A client-side bulk email sender for Gmail and Outlook, built with vanilla JavaScript. No backend required - all processing happens in your browser.

## Features

- **Data File Upload**: Parse CSV or Excel files (.xlsx, .xls) with email recipient data
- **Multiple Attachments**: Upload multiple attachment files (up to 20MB total recommended)
- **Dual Provider Support**: Send via Gmail API or Microsoft Graph API (Outlook)
- **OAuth Authentication**: Secure OAuth 2.0 authentication (no credentials stored)
- **Configurable Delays**: Set custom delays between sends with optional randomization
- **Real-Time Progress**: Track sending progress with detailed logging
- **Abort Capability**: Stop the sending process at any time
- **Anti-Spam Warnings**: Built-in warnings to help you stay compliant
- **100% Client-Side**: No backend, no database, no server-side processing

## Demo

[Live Demo](https://cj-1981.github.io/csv-email-sender/) (Coming Soon)

## Prerequisites

Before using this application, you need to:

1. **Have a Gmail or Outlook account** with sending permissions
2. **Configure OAuth credentials** for your chosen provider (see setup below)
3. **Understand email provider limits**:
   - Gmail: ~100-500 emails/day
   - Outlook: Similar limits
4. **Follow anti-spam best practices**:
   - Only email people who consented to receive messages
   - Include opt-out/unsubscribe information
   - Comply with CAN-SPAM Act and local laws

## Setup Instructions

### Step 1: Configure Gmail OAuth (Optional)

If you want to send via Gmail:

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the Gmail API:
   - Navigate to "APIs & Services" > "Library"
   - Search for "Gmail API" and enable it
4. Configure OAuth consent screen:
   - Navigate to "APIs & Services" > "OAuth consent screen"
   - Choose "External" user type
   - Fill in required fields (App name, User support email, Developer contact)
   - Add "https://www.googleapis.com/auth/gmail.send" as a scope
5. Create OAuth 2.0 credentials:
   - Navigate to "APIs & Services" > "Credentials"
   - Click "Create Credentials" > "OAuth client ID"
   - Application type: "Web application"
   - Name: "CSV Email Sender"
   - Authorized JavaScript origins: Add your domain without path (e.g., `https://username.github.io` or `http://localhost:8000` for local testing)
   - Click "Create"
6. Copy the **Client ID** (you'll need it in Step 3)

**Note:** With the new Google Identity Services, you don't need to configure Authorized redirect URIs.

### Step 2: Configure Outlook OAuth (Optional)

If you want to send via Outlook:

1. Go to [Azure Portal](https://portal.azure.com/)
2. Navigate to "Microsoft Entra ID (formerly Azure AD)" > "App registrations"
3. Click "New registration"
4. Fill in the form:
   - Name: "CSV Email Sender"
   - Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
   - Redirect URI: Web, enter your current page URL (e.g., `https://username.github.io/csv-email-sender/`)
   - Click "Register"
5. Configure API permissions:
   - Navigate to "API permissions" > "Add a permission"
   - Select "Microsoft Graph" > "Delegated permissions"
   - Search for and select "Mail.Send"
   - Click "Add permissions"
6. Copy the **Application (client) ID** (you'll need it in Step 3)

### Step 3: Configure OAuth Credentials in the App

The easiest way to configure your credentials is to use the built-in OAuth Setup Wizard:

1. Open the application in your browser
2. Click the "⚙️ Setup OAuth" button
3. Follow the step-by-step instructions to enter your Client ID(s)
4. Your credentials are stored in your browser's localStorage

**Alternative: Manual Configuration**
You can also configure credentials via browser console:
```javascript
// For Gmail
localStorage.setItem('oauth_gmail_client_id', 'YOUR_GMAIL_CLIENT_ID');

// For Outlook
localStorage.setItem('oauth_outlook_client_id', 'YOUR_OUTLOOK_CLIENT_ID');
```

### Step 4: Deploy to GitHub Pages

1. **Create a GitHub repository** (if you haven't already):
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   gh repo create csv-email-sender --public --source=.
   git push -u origin main
   ```

2. **Enable GitHub Pages**:
   - Go to your repository on GitHub
   - Navigate to "Settings" > "Pages"
   - Source: Deploy from a branch
   - Branch: `main` (or `master`) > `/ (root)`
   - Click "Save"

3. **Update OAuth redirect URIs** (if using your custom domain):
   - Go back to Google Cloud Console or Azure Portal
   - Update the authorized JavaScript origins and redirect URIs
   - Add your GitHub Pages URL: `https://username.github.io/csv-email-sender/`

4. **Access your application**:
   - Wait a few minutes for GitHub Pages to deploy
   - Visit `https://username.github.io/csv-email-sender/`

## Usage

### 1. Prepare Your Data File

You can download a CSV template directly from the app by clicking the "Download CSV Template" button in the upload section.

Or create a CSV or Excel file with the following columns:

**Supported File Formats:**
- **CSV** (.csv): Comma-separated values file
- **Excel** (.xlsx, .xls): Microsoft Excel workbook (reads first sheet only)

**Required Columns:**
- `recipient_email` (or `email`, `to`): Valid email address
- `subject`: Email subject line

**Optional Columns:**
- `attachment_filename` (or `attachment`): Filename matching uploaded attachment
- `body_content` (or `body`, `message`): Email body (overrides default body)

**Example CSV:**
```csv
recipient_email,subject,attachment_filename,body_content
user1@example.com,Welcome to our newsletter,newsletter.pdf,Thank you for subscribing!
user2@example.com,Monthly Update,report.pdf,Here is your monthly report.
user3@example.com,Special Offer,,Don't miss our special offer!
```

**Example Excel:**
- Create an Excel file (.xlsx or .xls)
- First row should contain column headers
- Fill in your data starting from row 2
- The first sheet in the workbook will be used

### 2. Upload Files

1. **Upload Data File**: Click "Choose File" under Data File and select your CSV or Excel file
2. **Upload Attachments**: Click "Choose File" under Attachment Files and select one or more files

### 3. Select Provider

1. Choose "Gmail" or "Outlook" from the dropdown
2. Click "Authenticate" button
3. Complete the OAuth flow in the popup window
4. Grant email sending permissions

### 4. Configure Sending

1. **Set Delay**: Enter delay between sends in milliseconds (recommended: 5000ms = 5 seconds)
2. **Enable Randomization**: Optional - adds random variation (±20%) to delays
3. **Set Default Body**: Optional - default email body for all emails

### 5. Send Emails

1. Click "Start Sending"
2. Monitor progress in real-time
3. Click "Abort" if needed to stop the process
4. View summary when complete

## Data File Format Examples

### CSV Format Examples

#### Simple Email List
```csv
recipient_email,subject
john@example.com,Hello World
jane@example.com,Meeting Reminder
```

#### With Custom Body Content
```csv
recipient_email,subject,body_content
john@example.com,Welcome,Welcome to our service!
jane@example.com,Update,Here is your monthly update.
```

#### With Attachments
```csv
recipient_email,subject,attachment_filename
john@example.com,Your Report,report.pdf
jane@example.com,Newsletter,newsletter.pdf
```

### Excel Format Instructions

You can also use Excel files (.xlsx or .xls) instead of CSV:

1. Open Microsoft Excel or your preferred spreadsheet application
2. Create a new workbook
3. Add column headers in the first row:
   - Row 1, Column A: `recipient_email`
   - Row 1, Column B: `subject`
   - Row 1, Column C: `attachment_filename` (optional)
   - Row 1, Column D: `body_content` (optional)
4. Fill in your data starting from row 2
5. Save the file as .xlsx or .xls format
6. Upload the file to the CSV Email Sender application

**Note:** Only the first sheet in your Excel workbook will be read.

## Excel Support

The CSV Email Sender now supports Excel files in addition to CSV files!

### Supported Excel Formats

- **.xlsx** - Excel Workbook (modern format)
- **.xls** - Excel 97-2003 Workbook (legacy format)

### How It Works

1. Upload your Excel file using the "Data File" upload field
2. The application automatically detects the file type
3. Excel files are parsed using the SheetJS library
4. Data from the first sheet is extracted and processed
5. The same validation rules apply as CSV files

### Excel File Requirements

- First row must contain column headers
- Required columns: `recipient_email`, `subject`
- Optional columns: `attachment_filename`, `body_content`
- Column names are case-insensitive
- Only the first sheet in the workbook is processed
- Protected workbooks are not supported

### Advantages of Using Excel

- **Easier Data Entry**: Excel's grid interface makes data entry simpler
- **Data Validation**: Use Excel's built-in data validation features
- **Formatting**: Apply formatting to ensure data consistency
- **Multiple Sheets**: Organize your data across sheets (only first sheet is read)
- **Formulas**: Use Excel formulas to generate content dynamically

### Tips for Excel Files

- Remove any empty rows before the header row
- Ensure there are no merged cells in the header row
- Save in .xlsx format for best compatibility
- Test with a small batch first before large sends

## Security Considerations

### What's Secure

- **No Backend**: No server-side processing or data storage
- **No Credential Storage**: OAuth tokens are kept in browser memory only
- **HTTPS Enforced**: GitHub Pages enforces HTTPS for all traffic
- **Limited Scope**: Only requests email sending permission (not full account access)
- **Input Validation**: All user inputs are validated and sanitized

### What You Need to Know

- **OAuth Credentials**: Your Client ID is public (this is safe for OAuth flows)
- **Token Lifetime**: Access tokens expire after 1 hour
- **No Persistence**: All data is cleared when you refresh the page
- **Your Responsibility**: You are responsible for your email compliance

## Anti-Spam Best Practices

To avoid account suspension and comply with anti-spam laws:

1. **Get Consent**: Only email people who explicitly consented to receive messages from you
2. **Provide Opt-Out**: Include unsubscribe information in your emails
3. **Identify Yourself**: Include your physical mailing address in emails
4. **Use Delays**: Always use delays between sends (minimum 1000ms recommended)
5. **Respect Limits**: Stay within your email provider's daily sending limits
6. **Monitor Bounces**: Remove invalid email addresses from your list
7. **Comply with Laws**: Follow CAN-SPAM Act (US), GDPR (EU), and local regulations

## Troubleshooting

### "Error 400: invalid_request" or "redirect_uri=storagerelay://file"

**Cause**: You're trying to use Gmail OAuth while opening the HTML file directly (file:// protocol).

**Solution**: You must run a local web server. Gmail OAuth does not work with file:// protocol.

Run one of these commands in your project directory:
```bash
# Python 3
python -m http.server 8000

# Node.js
npx serve

# PHP
php -S localhost:8000
```

Then open `http://localhost:8000` in your browser instead of opening the HTML file directly.

### "Authentication failed: Client ID not configured"

**Solution**: Configure OAuth credentials in `app.js` (see Setup Instructions above)

### "Missing required columns: recipient_email, subject"

**Solution**: Ensure your data file (CSV or Excel) has the required columns. Column names are case-insensitive but must match exactly (or use variations: `email`, `to` for recipient; `subject` is required)

### "Rate limit exceeded - Daily quota exhausted"

**Solution**: You've hit your email provider's daily sending limit. Wait until the limit resets (usually 24 hours)

### "Authentication failed - Token may have expired"

**Solution**: Re-authenticate by clicking the Authenticate button again. Tokens expire after 1 hour

### "Invalid recipient address"

**Solution**: Check that email addresses in your CSV are valid (format: user@domain.com)

### File not parsing correctly

**CSV Files**: Ensure your CSV is:
- Saved in UTF-8 encoding
- Uses comma as delimiter
- Has headers in the first row
- No special characters in column names

**Excel Files**: Ensure your Excel file:
- Is in .xlsx or .xls format (not .xlsm or other formats)
- Has headers in the first row of the first sheet
- Is not password-protected
- Contains valid data (no merged cells in header row)

## Browser Compatibility

- Chrome 90+
- Firefox 88+
- Safari 14+
- Edge 90+

## Performance

- **CSV Parsing**: < 2 seconds for 1,000 rows
- **Memory Usage**: Depends on CSV size and attachments
- **Recommended CSV Size**: < 10,000 rows
- **Recommended Attachment Size**: < 20MB total

## Limitations

- **No Batch Processing**: Emails are sent one-by-one to avoid rate limits
- **No Retry Logic**: Failed sends are logged but not automatically retried
- **No Template Engine**: Static body content only (no mail merge)
- **No Cross-Session Resume**: Progress is not saved between sessions
- **Attachment Size**: Limited by browser memory and provider limits (typically 25MB)

## Development

### Local Development

To run the application locally:

1. Clone the repository:
   ```bash
   git clone https://github.com/CJ-1981/csv-email-sender.git
   cd csv-email-sender
   ```

2. **IMPORTANT: You must use a local web server!**
   Gmail OAuth does NOT work when opening HTML files directly (file:// protocol). You must run a local web server.

   **Option A: Use the provided scripts (Easiest)**
   - Windows: Double-click `start-server.bat`
   - Mac/Linux: Run `bash start-server.sh` or `./start-server.sh`

   **Option B: Manual server startup**

   Choose one of these options:

   ```bash
   # Using Python 3 (recommended - comes with Python)
   python -m http.server 8000

   # Using Python 2
   python -m SimpleHTTPServer 8000

   # Using Node.js (with npx - no installation needed)
   npx serve

   # Using PHP
   php -S localhost:8000
   ```

3. Open your browser to `http://localhost:8000`

4. **Important**: Update OAuth redirect URIs to include `http://localhost:8000`

1. Clone the repository:
   ```bash
   git clone https://github.com/CJ-1981/csv-email-sender.git
   cd csv-email-sender
   ```

2. **IMPORTANT: You must use a local web server!**
   Gmail OAuth does NOT work when opening HTML files directly (file:// protocol). You must run a local web server.

   Choose one of these options:

   ```bash
   # Using Python 3 (recommended - comes with Python)
   python -m http.server 8000

   # Using Python 2
   python -m SimpleHTTPServer 8000

   # Using Node.js (with npx - no installation needed)
   npx serve

   # Using Node.js (with serve package installed)
   npm install -g serve
   serve

   # Using PHP
   php -S localhost:8000
   ```

3. Open your browser to `http://localhost:8000`

4. **Important**: Update OAuth redirect URIs to include `http://localhost:8000`
   - For Gmail: Add `http://localhost:8000` to Authorized JavaScript origins in Google Cloud Console
   - For Outlook: Add `http://localhost:8000` as a Redirect URI in Azure Portal

**Why can't I just open index.html directly?**

Google's OAuth 2.0 implementation requires a proper HTTP/HTTPS server and does not support the `file://` protocol. This is a security requirement from Google. If you try to authenticate while opening the file directly, you'll get a `redirect_uri=mismatch` or `invalid_request` error.

**Outlook authentication** may work without a server, but Gmail definitely requires one.

### File Structure

```
csv-email-sender/
├── index.html          # Main HTML structure
├── app.js              # Core application logic
├── styles.css          # Application styling
├── README.md           # This file
└── .moai/              # MoAI-ADK configuration
```

### CDN Libraries Used

- [Papa Parse](https://www.papaparse.com/) (v5.4.1) - CSV parsing
- [SheetJS](https://sheetjs.com/) (v0.20.1) - Excel parsing (Community Edition)
- [Google API Client](https://github.com/google/google-api-javascript-client) - Gmail API
- [Microsoft Graph SDK](https://github.com/microsoftgraph/msgraph-sdk-javascript) - Outlook API
- [MSAL.js](https://github.com/AzureAD/microsoft-authentication-library-for-js) - Microsoft Authentication

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Disclaimer

This tool is provided as-is for educational and personal use. Users are responsible for:

- Complying with all applicable laws and regulations (CAN-SPAM, GDPR, etc.)
- Obtaining proper consent before sending emails
- Respecting email provider terms of service and sending limits
- Any consequences of using this tool (including account suspension)

The authors are not responsible for any misuse of this tool.

## Support

If you encounter issues:

1. Check the [Troubleshooting](#troubleshooting) section
2. Review the [Setup Instructions](#setup-instructions)
3. [Open an issue](https://github.com/CJ-1981/csv-email-sender/issues) on GitHub

## Acknowledgments

- Built with vanilla JavaScript (no frameworks)
- Uses [Papa Parse](https://www.papaparse.com/) for CSV parsing
- Uses [SheetJS](https://sheetjs.com/) for Excel parsing
- Powered by Google Gmail API and Microsoft Graph API
- Hosted on GitHub Pages

---

**Made with ❤️ by CJ-1981**
