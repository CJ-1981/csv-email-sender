# CSV Email Sender

A client-side bulk email sender for Gmail and Outlook, built with vanilla JavaScript. No backend required - all processing happens in your browser.

## Features

- **CSV Upload**: Parse CSV files with email recipient data using Papa Parse
- **Multiple Attachments**: Upload multiple attachment files (up to 20MB total recommended)
- **Dual Provider Support**: Send via Gmail API or Microsoft Graph API (Outlook)
- **OAuth Authentication**: Secure OAuth 2.0 authentication (no credentials stored)
- **Configurable Delays**: Set custom delays between sends with optional randomization
- **Real-Time Progress**: Track sending progress with detailed logging
- **Abort Capability**: Stop the sending process at any time
- **Anti-Spam Warnings**: Built-in warnings to help you stay compliant
- **100% Client-Side**: No backend, no database, no server-side processing

## Demo

[Live Demo](https://chimin123.github.io/csv-email-sender/) (Coming Soon)

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
4. Create OAuth 2.0 credentials:
   - Navigate to "APIs & Services" > "Credentials"
   - Click "Create Credentials" > "OAuth client ID"
   - Application type: "Web application"
   - Name: "CSV Email Sender"
   - Authorized JavaScript origins: Add your GitHub Pages URL (e.g., `https://username.github.io`)
   - Authorized redirect URIs: Add your GitHub Pages URL
   - Click "Create"
5. Copy the **Client ID** (you'll need it in Step 3)

### Step 2: Configure Outlook OAuth (Optional)

If you want to send via Outlook:

1. Go to [Azure Portal](https://portal.azure.com/)
2. Navigate to "Azure Active Directory" > "App registrations"
3. Click "New registration"
4. Fill in the form:
   - Name: "CSV Email Sender"
   - Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
   - Redirect URI: Web, enter your GitHub Pages URL (e.g., `https://username.github.io/csv-email-sender/`)
   - Click "Register"
5. Configure API permissions:
   - Navigate to "API permissions" > "Add a permission"
   - Select "Microsoft Graph" > "Delegated permissions"
   - Search for and select "Mail.Send"
   - Click "Add permissions"
6. Copy the **Application (client) ID** (you'll need it in Step 3)

### Step 3: Update OAuth Credentials in app.js

1. Open `app.js` in a text editor
2. Find the `OAUTH_CONFIG` object at the top of the file
3. Replace the placeholder values with your actual credentials:

```javascript
const OAUTH_CONFIG = {
    gmail: {
        clientId: 'YOUR_GMAIL_CLIENT_ID_HERE', // Replace with your Gmail Client ID
        scopes: 'https://www.googleapis.com/auth/gmail.send',
        discoveryDocs: ['https://www.googleapis.com/discovery/v1/apis/gmail/v1/rest']
    },
    outlook: {
        clientId: 'YOUR_OUTLOOK_CLIENT_ID_HERE', // Replace with your Outlook Client ID
        scopes: ['Mail.Send'],
        authority: 'https://login.microsoftonline.com/common'
    }
};
```

4. Save the file

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

### 1. Prepare Your CSV File

Create a CSV file with the following columns:

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

### 2. Upload Files

1. **Upload CSV**: Click "Choose File" under CSV File and select your CSV file
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

## CSV Format Examples

### Simple Email List
```csv
recipient_email,subject
john@example.com,Hello World
jane@example.com,Meeting Reminder
```

### With Custom Body Content
```csv
recipient_email,subject,body_content
john@example.com,Welcome,Welcome to our service!
jane@example.com,Update,Here is your monthly update.
```

### With Attachments
```csv
recipient_email,subject,attachment_filename
john@example.com,Your Report,report.pdf
jane@example.com,Newsletter,newsletter.pdf
```

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

### "Authentication failed: Client ID not configured"

**Solution**: Configure OAuth credentials in `app.js` (see Setup Instructions above)

### "Missing required columns: recipient_email, subject"

**Solution**: Ensure your CSV has the required columns. Column names are case-insensitive but must match exactly (or use variations: `email`, `to` for recipient; `subject` is required)

### "Rate limit exceeded - Daily quota exhausted"

**Solution**: You've hit your email provider's daily sending limit. Wait until the limit resets (usually 24 hours)

### "Authentication failed - Token may have expired"

**Solution**: Re-authenticate by clicking the Authenticate button again. Tokens expire after 1 hour

### "Invalid recipient address"

**Solution**: Check that email addresses in your CSV are valid (format: user@domain.com)

### CSV not parsing correctly

**Solution**: Ensure your CSV is:
- Saved in UTF-8 encoding
- Uses comma as delimiter
- Has headers in the first row
- No special characters in column names

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
   git clone https://github.com/chimin123/csv-email-sender.git
   cd csv-email-sender
   ```

2. Serve the files using a local server:
   ```bash
   # Using Python 3
   python -m http.server 8000

   # Using Node.js (with npx)
   npx serve

   # Using PHP
   php -S localhost:8000
   ```

3. Open your browser to `http://localhost:8000`

4. **Important**: Update OAuth redirect URIs to include `http://localhost:8000`

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
3. [Open an issue](https://github.com/chimin123/csv-email-sender/issues) on GitHub

## Acknowledgments

- Built with vanilla JavaScript (no frameworks)
- Uses [Papa Parse](https://www.papaparse.com/) for CSV parsing
- Powered by Google Gmail API and Microsoft Graph API
- Hosted on GitHub Pages

---

**Made with ❤️ by chimin123**
