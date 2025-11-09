# Time Tracker - Deployment Instructions

## Overview

This document explains how to deploy the Time Tracker web app with the correct permissions settings to allow all users to use the app **without needing direct edit access to the Google Sheet**.

## The Problem

Previously, users needed direct edit permissions on the Google Sheet to use the Time Tracker app. This created a security concern because the sheet should only be editable by admins and super admins.

## The Solution

We've updated the code to use `Session.getEffectiveUser()` instead of `Session.getActiveUser()`. This allows the script to:
- Run with the script owner's permissions (who should be an admin with edit access)
- Identify the accessing user correctly
- Write to the sheet on behalf of users without giving them direct sheet access

## Deployment Steps

### 1. Open the Apps Script Editor

1. Open your Google Sheet for the Time Tracker
2. Click **Extensions** → **Apps Script**
3. Ensure your `code.gs` and `Index.html` files are up to date with the latest changes

### 2. Deploy the Web App

1. In the Apps Script editor, click **Deploy** → **New deployment**
2. Click the gear icon ⚙️ next to "Select type" and choose **Web app**
3. Configure the deployment settings:

   **Important Settings:**

   - **Description**: Enter a description like "Time Tracker - Service Account Mode"

   - **Execute as**: **Me** ← **CRITICAL: Select "Me" (your email address)**
     - This makes the script run with YOUR permissions (the script owner)
     - Since you're an admin, the script will have edit access to the sheet
     - Users won't need their own edit permissions

   - **Who has access**: Choose one of the following based on your needs:
     - **Only myself** - Only you can access (for testing)
     - **Anyone within [your organization]** - Anyone in your Google Workspace domain ← **RECOMMENDED**
     - **Anyone** - Anyone with the link (not recommended for internal apps)

4. Click **Deploy**
5. **Grant permissions** when prompted:
   - Click **Authorize access**
   - Select your admin account
   - Click **Advanced** if you see a warning
   - Click **Go to [Project name] (unsafe)** (this is safe - it's your own script)
   - Review the permissions and click **Allow**

6. Copy the **Web app URL** that appears after deployment
7. Update the `WEB_APP_URL` constant in `code.gs` (line 25) with this URL

### 3. Update the Spreadsheet Permissions

Now that the app runs with your (admin) permissions, you can restrict sheet access:

1. Open the Google Sheet
2. Click the **Share** button
3. Set permissions as follows:
   - **Super Admins**: Editor access (e.g., cantiporta@threecolts.com)
   - **Admins**: Editor access (e.g., apelagio@threecolts.com)
   - **Regular Users**: **No direct access** (remove them from the share settings)

4. Make sure "General access" is set to **Restricted** (only people with access can open)

### 4. Share the Web App URL with Users

Users can now access the Time Tracker app via the Web App URL without needing sheet permissions:

```
https://script.google.com/a/macros/threecolts.com/s/[YOUR_DEPLOYMENT_ID]/exec
```

Users will:
- Open the web app URL
- Be prompted to sign in with their Google Workspace account
- See the Time Tracker interface
- Be able to clock in/out, take breaks, etc.
- All actions will be written to the sheet using your (the owner's) permissions

## How It Works

### Before (Session.getActiveUser)
```javascript
const userEmail = Session.getActiveUser().getEmail();
// When running as user: Returns the user's email ✓
// When running as owner: Returns the owner's email ✗
```

### After (Session.getEffectiveUser)
```javascript
const userEmail = Session.getEffectiveUser().getEmail();
// When running as user: Returns the user's email ✓
// When running as owner: Returns the accessing user's email ✓
```

### Permission Flow

```
User opens web app URL
    ↓
Google authenticates user (user@domain.com)
    ↓
Script runs as owner (admin@domain.com) with owner's permissions
    ↓
Session.getEffectiveUser() returns user@domain.com
    ↓
Script writes to sheet using owner's edit permissions
    ↓
User sees their data in the app interface
```

## Testing the Deployment

### Test as Admin
1. Open the web app URL while logged in as an admin
2. Verify you can clock in/out and see admin features

### Test as Regular User
1. Have a regular user (non-admin) open the web app URL
2. Verify they can clock in/out
3. Verify they CANNOT directly edit the Google Sheet
4. Check that their time entries appear correctly in the sheet

### Test Permissions
1. Try to access the sheet directly as a regular user - should be denied
2. Try to use the app as a regular user - should work
3. Verify admin-only features are hidden from regular users

## Troubleshooting

### "Permission denied" errors
- Ensure the deployment is set to "Execute as: Me"
- Verify the script owner (you) has edit access to the sheet
- Re-authorize the script permissions

### Users can't access the app
- Check "Who has access" setting in deployment
- Verify users are in your Google Workspace domain (if using domain restriction)
- Make sure users are using the correct web app URL

### Calendar integration not working
- The script needs Calendar API enabled in your Google Cloud project
- Go to **Extensions** → **Apps Script** → **Project Settings**
- Under "Google Cloud Platform (GCP) Project", ensure Calendar API is enabled
- Users may need to grant calendar access when first using the app

### Telegram notifications not working
- Verify the `TELEGRAM_BOT_TOKEN` is correct in `code.gs`
- Check that users have linked their Telegram chat IDs via the admin panel

## Security Considerations

### What This Approach Protects
✅ Sheet data is only editable by admins directly
✅ Users can't modify formulas or sheet structure
✅ Users can't see other users' data in the sheet directly
✅ All writes are controlled by your Apps Script logic

### What Users Can Still Do
- View their own time entries through the app
- Clock in/out through the app interface
- View their calendar events (requires their consent)
- Upload/delete their own profile pictures and backgrounds

### Admin Controls
Admins (defined in `SUPER_ADMIN_EMAILS` and `ADMIN_EMAILS` in `code.gs`) can:
- Access admin panel in the app
- Generate new week sheets
- View all users' status
- Modify the sheet directly
- View all users' background images

## Updating the Deployment

If you make changes to the code:

1. Save your changes in the Apps Script editor
2. Click **Deploy** → **Manage deployments**
3. Click the edit icon (pencil) next to your active deployment
4. Under "Version", select **New version**
5. Add a description of the changes
6. Click **Deploy**
7. The web app URL remains the same - users don't need to update anything

## Support

For issues or questions:
- Check the Apps Script execution logs: **Executions** tab in Apps Script editor
- Review error messages in the browser console (F12)
- Verify all environment settings match this guide
- Contact the development team

---

**Last Updated**: 2025-11-09
**Version**: 2.0 (Service Account Mode)
