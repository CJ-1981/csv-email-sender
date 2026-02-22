# Quick Start Guide - CSV Email Sender

## 🚀 5-Minute Setup

### Step 1: Configure OAuth (Required)

Before using the application, you MUST configure OAuth credentials:

#### Option A: Gmail Only (Recommended for Beginners)

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project
3. Enable "Gmail API" (APIs & Services > Library)
4. Create OAuth 2.0 Client ID:
   - Application type: "Web application"
   - Name: "CSV Email Sender"
   - Authorized origins: `http://localhost:8000` (for testing) or your GitHub Pages URL
5. Copy the **Client ID**

#### Option B: Outlook Only

1. Go to [Azure Portal](https://portal.azure.com/)
2. Register a new app in Azure Active Directory
3. Add "Mail.Send" permission from Microsoft Graph
4. Copy the **Application (client) ID**

#### Option C: Both Providers

Complete both Option A and Option B above.

### Step 2: Update app.js

1. Open `app.js` in a text editor
2. Find the `OAUTH_CONFIG` section at the top
3. Replace the placeholder values:

```javascript
const OAUTH_CONFIG = {
    gmail: {
        clientId: 'PASTE_YOUR_GMAIL_CLIENT_ID_HERE',  // Replace this
        // ... keep other settings
    },
    outlook: {
        clientId: 'PASTE_YOUR_OUTLOOK_CLIENT_ID_HERE', // Replace this
        // ... keep other settings
    }
};
```

4. Save the file

### Step 3: Run Locally (Testing)

```bash
# Option 1: Python
python -m http.server 8000

# Option 2: Node.js
npx serve

# Option 3: PHP
php -S localhost:8000
```

Then open: http://localhost:8000

### Step 4: Deploy to GitHub Pages (Production)

```bash
# Initialize git repo
git init
git add .
git commit -m "Initial commit"

# Create GitHub repo and push
gh repo create csv-email-sender --public --source=.
git push -u origin main
```

Then enable GitHub Pages in repository Settings.

---

## 📝 Create Your CSV

### Minimal CSV (2 columns required)

```csv
recipient_email,subject
john@example.com,Test Email
jane@example.com,Hello World
```

### Full CSV (with attachments and custom body)

```csv
recipient_email,subject,attachment_filename,body_content
john@example.com,Your Report,report.pdf,Please find attached report.
jane@example.com,Newsletter,newsletter.pdf,Monthly newsletter inside.
bob@example.com,Special Offer,,Don't miss this offer!
```

**Column Names (Case-Insensitive)**:
- **Required**: `recipient_email` (or `email`, `to`), `subject`
- **Optional**: `attachment_filename` (or `attachment`), `body_content` (or `body`, `message`)

---

## 🎯 Basic Usage

1. **Upload CSV**: Click "Choose File" → Select your CSV
2. **Upload Attachments** (optional): Click "Choose File" → Select files
3. **Select Provider**: Choose "Gmail" or "Outlook"
4. **Authenticate**: Click "Authenticate" → Complete OAuth flow
5. **Configure Delay**: Set delay (recommended: 5000ms = 5 seconds)
6. **Start Sending**: Click "Start Sending" → Confirm → Monitor progress

---

## ⚠️ Important Warnings

### Daily Sending Limits
- **Gmail**: 100-500 emails/day
- **Outlook**: Similar limits

### Best Practices
1. ✅ Use delays between sends (minimum 1000ms)
2. ✅ Only email people who consented
3. ✅ Include unsubscribe information
4. ✅ Comply with CAN-SPAM Act
5. ❌ Don't send spam
6. ❌ Don't use zero delays

### Risk of Account Suspension
- Spam-like behavior may result in account suspension
- Start with small batches (10-20 emails)
- Monitor for warnings from email provider

---

## 🔧 Troubleshooting

### "Client ID not configured"
→ Configure OAuth credentials in app.js (see Step 2 above)

### "Missing required columns"
→ Check your CSV has `recipient_email` and `subject` columns

### "Rate limit exceeded"
→ You've hit daily quota. Wait 24 hours.

### Emails not sending
→ Check:
- OAuth authentication successful
- Email addresses are valid
- Network connection is stable
- Browser console for errors (F12)

---

## 📚 More Information

- **Full Documentation**: See [README.md](README.md)
- **Setup Guide**: See README.md "Setup Instructions" section
- **CSV Examples**: See README.md "CSV Format Examples" section
- **Troubleshooting**: See README.md "Troubleshooting" section

---

## ✅ Pre-Flight Checklist

Before sending to large lists:

- [ ] OAuth credentials configured
- [ ] Tested with 3-5 emails
- [ ] CSV validated (all emails valid format)
- [ ] Attachments uploaded (if needed)
- [ ] Delay configured (minimum 1000ms)
- [ ] Understand daily sending limits
- [ ] Have consent from recipients
- [ ] Prepared for account suspension risk

---

**Need Help?** Check the [README.md](README.md) or [open an issue](https://github.com/chimin123/csv-email-sender/issues)
