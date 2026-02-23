# Outlook OAuth Setup Guide

## Common Error: AADSTS50194

**Error Message:**
```
Application 'xxx' is not configured as a multi-tenant application.
Usage of the /common endpoint is not supported for such applications
created after '10/15/2018'.
```

**Cause:** Your app is configured as single-tenant but the code uses the `/common` endpoint.

**Solution:** Follow the steps below to fix this.

---

## Option 1: Make App Multi-Tenant (Recommended)

This allows ANY Microsoft account to authenticate (personal Outlook, work, school).

### Steps:

1. **Go to Azure Portal**: https://portal.azure.com/

2. **Find your app registration**:
   - Search for "App registrations" in the search bar
   - Click on "csv-email-sender" (or your app name)

3. **Update Authentication settings**:
   - In the left menu, click **Authentication**
   - Find the **Supported account types** section
   - Click the edit button (pencil icon) or select the dropdown
   - Change from:
     ```
     Accounts in this organizational directory only
     ```
   - To:
     ```
     Accounts in any organizational directory and personal Microsoft accounts
     ```
   - Click **Save** (or **Configure**)

4. **Verify Redirect URI**:
   - Still in the Authentication section
   - Under **Redirect URIs**, make sure you have:
     - `https://cj-1981.github.io/csv-email-sender/`
     - (Optional) `http://localhost:8000` for local testing
   - Platform should be **Web**
   - Click **Save** if you made changes

5. **Wait for propagation**:
   - Azure changes typically take 1-2 minutes
   - Sometimes can take up to 5 minutes

6. **Test again**:
   - Go to: https://cj-1981.github.io/csv-email-sender/
   - Select "Outlook"
   - Click "Authenticate"
   - It should work now!

---

## Option 2: Keep Single-Tenant (Organization Only)

If you ONLY want users from YOUR organization to use the app:

### Understanding Single-Tenant vs Multi-Tenant

| Type | Who Can Authenticate | Use Case |
|------|---------------------|----------|
| **Single-Tenant** | Only accounts in your organization | Internal company apps |
| **Multi-Tenant** | Any Microsoft account (personal + work/school) | Public apps, personal use |

### Steps for Single-Tenant:

1. **Keep your app as single-tenant** (do nothing)

2. **Get your organization ID**:
   - In Azure Portal, go to **Microsoft Entra ID** (formerly Azure AD)
   - Click **Properties** in the left menu
   - Copy your **Tenant ID** (looks like: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)

3. **Update the app code**:
   - The app needs to be updated to use your tenant-specific endpoint
   - Instead of `/common`, it would use `/your-tenant-id`

**Note:** This requires code changes. For personal use or public apps, **Option 1 (Multi-Tenant)** is much simpler.

---

## Verification Checklist

Before testing, verify these settings in Azure Portal:

### App Registration → Authentication:
- [x] **Supported account types**: "Accounts in any organizational directory and personal Microsoft accounts"
- [x] **Redirect URIs**: Contains `https://cj-1981.github.io/csv-email-sender/`
- [x] **Platform**: Web

### App Registration → API Permissions:
- [x] **Microsoft Graph** → **Delegated permissions** → **Mail.Send**
- [x] Click **Grant admin consent** button (if you see it)

### App Registration → Overview:
- [x] Copy your **Application (client) ID** - you'll need this for the app
- [x] Make sure you've entered this in the OAuth Setup Wizard

---

## Troubleshooting

### Error: "AADSTS50194"
**Solution:** Follow Option 1 above to make app multi-tenant

### Error: "AADSTS650053"
**Cause:** Redirect URI not configured or doesn't match
**Solution:** Add your GitHub Pages URL to Redirect URIs in Azure Portal

### Error: "invalid_client"
**Cause:** Wrong Client ID
**Solution:** Copy the Application (client) ID from Azure Portal and re-enter in OAuth Setup Wizard

### Error: "interaction_required"
**Cause:** Need to grant admin consent
**Solution:**
1. Go to **API permissions** in Azure Portal
2. Click **Grant admin consent** button
3. Wait for permissions to propagate

---

## Local Development Setup

For testing locally:

1. In Azure Portal, under **Redirect URIs**, add:
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

4. Authenticate with Outlook

---

## Security Notes

- **Multi-tenant apps** can be used by anyone with a Microsoft account
- **Client ID is public** - it's safe to share (like a username)
- **Tokens are stored in browser memory only** - not saved on any server
- **OAuth tokens expire** - users must re-authenticate after expiration

---

## Need More Help?

- [Microsoft Identity Platform Documentation](https://docs.microsoft.com/azure/active-directory/develop/)
- [MSAL.js Documentation](https://github.com/AzureAD/microsoft-authentication-library-for-js)
- [Azure Portal Help](https://docs.microsoft.com/azure/)

---

## Quick Fix Summary

**If you see AADSTS50194 error:**

1. Go to Azure Portal → App Registrations
2. Select your app
3. Go to Authentication
4. Change "Supported account types" to "Accounts in any organizational directory and personal Microsoft accounts"
5. Save
6. Wait 2 minutes
7. Try again

That's it! 🎉
