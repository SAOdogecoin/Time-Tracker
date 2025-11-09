# Setting Up Domain-Wide Delegation for Time Tracker

## Prerequisites
✅ You are a Google Workspace Super Admin
✅ You created the Time Tracker app with your Gmail account
✅ The code changes have been deployed (Session.getEffectiveUser)

## Overview

This guide will walk you through enabling domain-wide delegation so that:
- Your app runs with YOUR permissions (can write to the sheet)
- But accesses EACH USER'S calendar (not just yours)
- Users don't need sheet edit permissions

**Time Required**: 15-20 minutes

---

## Part 1: Configure the Google Cloud Project

### Step 1: Open Your Apps Script Project

1. Open your Time Tracker Google Sheet
2. Click **Extensions** → **Apps Script**
3. You should see your `code.gs` and `Index.html` files

### Step 2: Check Your GCP Project

1. In Apps Script editor, click the **gear icon** ⚙️ (Project Settings) on the left
2. Scroll down to **Google Cloud Platform (GCP) Project**
3. You'll see one of two things:

   **Option A**: You see a project number (like `project-id-123456789`)
   - Note this project number
   - Click on it to open Google Cloud Console

   **Option B**: It says "Default project"
   - You need to create a standard GCP project
   - Click **Change project**
   - Click **Create a new Google Cloud project**
   - Wait 1-2 minutes for it to be created
   - Note the new project number

### Step 3: Enable Required APIs

1. You should now be in Google Cloud Console (console.cloud.google.com)
2. Make sure the correct project is selected in the top dropdown
3. Go to **APIs & Services** → **Library**
4. Search for and enable these APIs (click **ENABLE** for each):
   - ✅ **Google Calendar API**
   - ✅ **Google Sheets API** (probably already enabled)
   - ✅ **Google Drive API** (probably already enabled)

### Step 4: Configure OAuth Consent Screen

1. In Google Cloud Console, go to **APIs & Services** → **OAuth consent screen**
2. If you haven't set this up yet:
   - Select **Internal** (only users in your Workspace can access)
   - Click **CREATE**
3. Fill in the required fields:
   - **App name**: `Time Tracker`
   - **User support email**: Your email (cantiporta@threecolts.com)
   - **Developer contact**: Your email
   - Click **SAVE AND CONTINUE**
4. On the "Scopes" page, click **ADD OR REMOVE SCOPES**
5. Add these scopes:
   ```
   https://www.googleapis.com/auth/calendar.readonly
   https://www.googleapis.com/auth/spreadsheets
   https://www.googleapis.com/auth/drive.file
   https://www.googleapis.com/auth/script.external_request
   ```
   - You can paste these into the "Manually add scopes" box at the bottom
   - Click **ADD TO TABLE**
   - Click **UPDATE**
   - Click **SAVE AND CONTINUE**
6. Review the summary and click **BACK TO DASHBOARD**

---

## Part 2: Enable Domain-Wide Delegation

### Step 5: Get Your OAuth Client ID

1. Still in Google Cloud Console, go to **APIs & Services** → **Credentials**
2. You should see an OAuth 2.0 Client ID that was automatically created
   - It might be named something like "Apps Script" or "Google Apps Script"
3. Click on it to view details
4. **Copy the Client ID** - It's a long number like `123456789012-abc123def456.apps.googleusercontent.com`
   - Keep this handy - you'll need it in the next step

**Alternative method if you don't see an OAuth Client:**
1. Go back to Apps Script editor
2. Run any function (like `doGet`) to trigger authorization
3. Authorize the script when prompted
4. Go back to Cloud Console → Credentials
5. You should now see the OAuth 2.0 Client ID

### Step 6: Enable Domain-Wide Delegation in GCP

1. In Google Cloud Console, still in **Credentials**
2. Click on your OAuth 2.0 Client ID
3. Check the box **"Enable G Suite Domain-wide Delegation"** (or "Enable Google Workspace Domain-wide Delegation")
4. **Product name**: `Time Tracker`
5. Click **SAVE**
6. **Copy the Client ID again** (you'll need it for the next part)

---

## Part 3: Authorize in Google Workspace Admin Console

### Step 7: Add API Client to Domain-Wide Delegation

1. Open a new tab and go to [Google Admin Console](https://admin.google.com/)
2. Sign in with your Super Admin account (cantiporta@threecolts.com)
3. Navigate to **Security** → **Access and data control** → **API controls**
4. Scroll down to **Domain-wide delegation**
5. Click **MANAGE DOMAIN WIDE DELEGATION**
6. Click **Add new**
7. Fill in the form:
   - **Client ID**: Paste the Client ID from Step 5/6
   - **OAuth Scopes**: Add these scopes (comma-separated):
     ```
     https://www.googleapis.com/auth/calendar.readonly,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/drive.file
     ```
   - **Description** (optional): `Time Tracker App`
8. Click **AUTHORIZE**

**You should see a success message!** The app is now authorized for domain-wide delegation.

---

## Part 4: Deploy the Web App

### Step 8: Deploy with Correct Settings

1. Go back to your **Apps Script editor**
2. Click **Deploy** → **Manage deployments**
3. If you have an existing deployment:
   - Click the **Edit** icon (pencil) next to it
   - Select **New version**
4. If you don't have a deployment yet:
   - Click **Deploy** → **New deployment**
   - Click the gear icon and select **Web app**

5. Configure these settings:
   - **Description**: `Time Tracker - Domain-Wide Delegation Enabled`
   - **Execute as**: **Me (your-email@threecolts.com)** ← CRITICAL
   - **Who has access**: **Anyone within threecolts.com** (your organization)

6. Click **Deploy**
7. **Authorize** when prompted:
   - Click **Authorize access**
   - Select your account
   - Click **Advanced** if you see a warning
   - Click **Go to Time Tracker (unsafe)** (it's safe - it's your app)
   - Review permissions and click **Allow**

8. **Copy the Web App URL** - It looks like:
   ```
   https://script.google.com/a/macros/threecolts.com/s/AKfycb.../exec
   ```

### Step 9: Update WEB_APP_URL in Code (if needed)

1. In `code.gs`, find line 25:
   ```javascript
   const WEB_APP_URL = "https://script.google.com/...";
   ```
2. Update it with your new deployment URL if it changed
3. Click **Save**
4. **Redeploy** if you changed the URL:
   - **Deploy** → **Manage deployments**
   - Click **Edit** on your deployment
   - Select **New version**
   - Click **Deploy**

---

## Part 5: Configure Sheet Permissions

### Step 10: Lock Down the Sheet

Now that the app runs with your permissions, you can restrict direct sheet access:

1. Open your Time Tracker Google Sheet
2. Click the **Share** button (top right)
3. Set permissions:
   - **You (cantiporta@threecolts.com)**: Editor (owner)
   - **Admin (apelagio@threecolts.com)**: Editor
   - **All other users**: REMOVE them from the share list

4. Under "General access":
   - Set to **Restricted** (only people with access can open)

5. Click **Done**

**Important**: Users will now get "Permission denied" if they try to open the sheet directly. This is correct! They should use the web app URL instead.

---

## Part 6: Test the Setup

### Step 11: Test as Super Admin

1. Open the **Web App URL** in a browser
2. You should see the Time Tracker interface
3. Test these features:
   - ✅ Clock in/out (should write to sheet)
   - ✅ Calendar events show up (YOUR calendar)
   - ✅ Admin panel is visible

### Step 12: Test as Regular User

1. Open the **Web App URL** in an incognito window OR ask another user to test
2. Sign in as a regular user (not admin)
3. Test these features:
   - ✅ Clock in/out (should write to sheet)
   - ✅ Calendar events show up (THEIR calendar, not yours!)
   - ❌ Admin panel should be hidden
   - ❌ Can't open the sheet directly (should get permission denied)

4. Check the Google Sheet:
   - Open the sheet (you can as admin)
   - Verify the user's time entry was recorded
   - Verify it has their username, not yours

---

## Part 7: Troubleshooting

### Issue: Users See "Authorization Required" Dialog

**Cause**: The app needs to request calendar access from each user the first time

**Solution**: This is expected behavior! Each user will see this dialog once:
- They'll see a list of permissions (Calendar, Sheets, Drive)
- They click **Allow**
- They won't see it again (unless you add new permissions)

### Issue: Users See YOUR Calendar Events

**Possible causes**:
1. **Domain-wide delegation not authorized correctly**
   - Check Step 7 - make sure the Client ID and scopes are correct
   - Wait 10-15 minutes for changes to propagate

2. **Deployment settings wrong**
   - Make sure you deployed with "Execute as: Me"
   - Check Step 8

3. **Code still using wrong calendar API call**
   - Check line 514 in code.gs
   - Should be: `Calendar.Events.list(userEmail, {...})`
   - NOT: `Calendar.Events.list('primary', {...})`

### Issue: "Requested entity was not found" for Calendar

**Cause**: The script can't access the user's calendar with domain-wide delegation

**Solution**:
1. Verify the scope is correct in Admin Console (Step 7)
2. Make sure it's `.../auth/calendar.readonly` not `.../auth/calendar`
3. Wait 10-15 minutes for delegation to take effect
4. Have the user try again

### Issue: Sheet Writes Fail

**Cause**: Script is not running with your permissions

**Solution**:
- Check deployment is "Execute as: Me"
- Make sure YOU (the owner) have edit access to the sheet
- Redeploy the app

### Issue: "Access Denied" from Calendar API

**Cause**: The Calendar API might not be enabled or domain-wide delegation isn't working

**Solution**:
1. Verify Calendar API is enabled (Step 3)
2. Verify domain-wide delegation is authorized (Step 7)
3. Check that you used the correct Client ID
4. Wait 15 minutes and try again (propagation delay)

---

## Understanding What's Happening Behind the Scenes

### When a User Opens the App:

```
1. User opens: https://script.google.com/.../exec
2. Google authenticates: Who is this user?
   → Returns: user@threecolts.com
3. Script executes as: YOU (cantiporta@threecolts.com)
   → Script has YOUR permissions
4. Session.getEffectiveUser() returns: user@threecolts.com
   → Script knows which user is accessing
5. Calendar API call with domain-wide delegation:
   Calendar.Events.list(user@threecolts.com)
   → Script uses YOUR authority + domain delegation
   → Returns: user@threecolts.com's calendar (not yours!)
6. Sheet write:
   SpreadsheetApp.getActiveSpreadsheet()
   → Uses YOUR edit permissions
   → Writes to sheet successfully
```

### The Magic of Domain-Wide Delegation:

Without it:
- Script as YOU → Can only see YOUR calendar
- Script as User → Can't write to sheet

With it:
- Script as YOU → Can write to sheet ✅
- Script as YOU + Delegation → Can see ANY user's calendar ✅

---

## Security Notes

### What Users CAN Do:
✅ Use the web app to track time
✅ See their own calendar events through the app
✅ Upload their own profile pictures/backgrounds
✅ Delete their own backgrounds

### What Users CANNOT Do:
❌ Open the Google Sheet directly
❌ Edit the sheet directly
❌ Modify formulas or sheet structure
❌ See other users' data in the sheet
❌ Delete other users' backgrounds

### What Admins CAN Do:
✅ Everything users can do
✅ Open and edit the sheet directly
✅ Generate new week sheets
✅ View all users' status
✅ Access admin panel in the app

---

## Summary Checklist

Before going live, verify:

- [ ] Google Cloud Project has Calendar, Sheets, and Drive APIs enabled
- [ ] OAuth consent screen is configured with correct scopes
- [ ] Domain-wide delegation is enabled in GCP for your OAuth client
- [ ] Client ID is authorized in Google Workspace Admin Console with correct scopes
- [ ] Web app is deployed with "Execute as: Me" and correct access level
- [ ] WEB_APP_URL in code.gs matches your deployment URL
- [ ] Sheet permissions are restricted (only admins have direct access)
- [ ] Tested as super admin - works correctly
- [ ] Tested as regular user - sees their own calendar and can write time entries
- [ ] Regular user CANNOT open sheet directly

---

## Need Help?

- **Google's Official Guide**: [Domain-Wide Delegation](https://developers.google.com/identity/protocols/oauth2/service-account#delegatingauthority)
- **Apps Script Documentation**: [Understanding Authorization](https://developers.google.com/apps-script/guides/services/authorization)
- **Check Execution Logs**: Apps Script editor → **Executions** tab

---

**Last Updated**: 2025-11-09
**Your Role**: Super Admin & App Creator
**Status**: Ready to implement ✅
