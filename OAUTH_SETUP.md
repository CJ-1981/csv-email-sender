# Gmail OAuth Setup Guide

## Quick Setup: Add Test Users (Recommended for Personal Use)

### Step 1: Go to Google Cloud Console
Visit: https://console.cloud.google.com/

### Step 2: Select Your Project
- Make sure you selected the correct project for CSV Email Sender

### Step 3: Go to OAuth Consent Screen
- Navigate to: **APIs & Services** > **OAuth consent screen**
- You should see your app configuration

### Step 4: Add Test Users
1. Scroll down to the **Test users** section
2. Click the **+ Add users** button
3. Enter your email address: `JungChimin@gmail.com`
4. You can add multiple emails if needed (comma-separated)
5. Click **Add**
6. Click **Save** at the bottom of the page

### Step 5: Wait and Test
- Changes take effect immediately
- Go back to: https://cj-1981.github.io/csv-email-sender/
- Try authenticating again - it should work now!

### Step 6: Verify Your Authorized Origins
While in Google Cloud Console:

1. Navigate to **APIs & Services** > **Credentials**
2. Click on your OAuth 2.0 Client ID
3. Under **Authorized JavaScript origins**, make sure you have:
   ```
   https://cj-1981.github.io
   ```
4. **Important**: Remove any `file://` or local development URLs from production
5. Click **Save**

---

## Alternative: Publish for Production Use

**Only do this if you want public access to your app.**

### Why This Requires More Work:
- Google requires app verification for sensitive scopes (like `gmail.send`)
- Verification costs $75 USD
- Process takes several days
- You'll need to provide:
  - Domain ownership proof
  - Privacy policy URL
  - Terms of service URL
  - Organization information

### Steps to Publish:
1. Go to **APIs & Services** > **OAuth consent screen**
2. Change **Publishing status** from **Testing** to **In production**
3. Follow Google's verification process

**Recommendation**: Use test users (Option above) for personal/internal use.

---

## Local Development Setup

If you want to test locally:

1. In Google Cloud Console, under **Authorized JavaScript origins**, add:
   ```
   http://localhost:8000
   ```

2. Start a local server:
   ```bash
   # Windows: Double-click start-server.bat
   # Mac/Linux: ./start-server.sh
   # Or manually: python -m http.server 8000
   ```

3. Open: http://localhost:8000

4. Add your test user as shown in Steps 1-5 above

---

## Troubleshooting

### Error: "Access blocked: has not completed verification"
**Solution**: Add your email as a test user (see Steps 1-5 above)

### Error: "redirect_uri_mismatch"
**Solution**: Make sure `https://cj-1981.github.io` is in Authorized JavaScript origins

### Error: "unauthorized_client"
**Solution**: Check that the Client ID in the OAuth Setup Wizard matches your Google Cloud Console

### Can I add more test users?
**Yes**: You can add up to 100 test users in Testing mode

---

## Security Notes

- **Test users** can only access the app while it's in Testing mode
- **Testing mode** expires after 6 months - you'll need to publish or extend it
- Never share your Client ID publicly (though it's not a secret)
- Your OAuth tokens are stored in browser memory only, not on any server

---

## Need More Help?

- [Google OAuth 2.0 Documentation](https://developers.google.com/identity/protocols/oauth2)
- [Google Cloud Console Help](https://cloud.google.com/docs/help)
- [Gmail API Documentation](https://developers.google.com/gmail/api)
