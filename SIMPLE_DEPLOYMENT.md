# Simple Deployment Guide - Time Tracker (No Calendar)

## ✅ Problem Solved!

Users can now use the Time Tracker app to clock in/out **without needing edit access to the Google Sheet**.

**Calendar feature has been disabled** to avoid the need for Google Workspace Admin Console access (domain-wide delegation).

---

## What This Deployment Does

✅ **Sheet writes work** - Script runs with your (owner's) permissions
✅ **No user sheet permissions needed** - Users only access via web app
✅ **Secure** - Sheet structure protected, users can't edit directly
❌ **Calendar disabled** - No calendar events shown (can be re-enabled later with admin access)

---

## Quick Deployment (5 Minutes)

### Step 1: Deploy the Web App

1. Open your Time Tracker Google Sheet
2. Click **Extensions** → **Apps Script**
3. Verify your `code.gs` has the latest changes (Session.getEffectiveUser)
4. Click **Deploy** → **New deployment** (or **Manage deployments** if updating)
5. Click the gear icon ⚙️ and select **Web app**

### Step 2: Configure Deployment Settings

**CRITICAL SETTINGS:**

- **Description**: `Time Tracker - No Calendar Mode`
- **Execute as**: **Me (your-email@threecolts.com)** ← MUST BE "Me"
- **Who has access**: **Anyone within threecolts.com** (or "Anyone" if needed)

### Step 3: Authorize and Deploy

1. Click **Deploy**
2. When prompted, click **Authorize access**
3. Select your account
4. If you see a warning "This app isn't verified":
   - Click **Advanced**
   - Click **Go to Time Tracker (unsafe)** (it's safe - it's YOUR app)
5. Review permissions and click **Allow**
6. **Copy the Web App URL** that appears

### Step 4: Update WEB_APP_URL (If Needed)

1. In `code.gs`, find line 25:
   ```javascript
   const WEB_APP_URL = "https://script.google.com/...";
   ```
2. If the URL changed, update it with your new deployment URL
3. Click **Save** (Ctrl+S or Cmd+S)
4. If you changed it, **redeploy**:
   - **Deploy** → **Manage deployments**
   - Click **Edit** (pencil icon)
   - Select **New version**
   - Click **Deploy**

---

## Step 5: Restrict Sheet Permissions

Now that the app runs with your permissions, lock down the sheet:

1. Open your Time Tracker Google Sheet
2. Click **Share** button (top right)
3. Set these permissions:
   - **You (cantiporta@threecolts.com)**: Editor (owner)
   - **Other Admins (apelagio@threecolts.com)**: Editor
   - **All regular users**: **REMOVE** from share list
4. Set "General access" to **Restricted**
5. Click **Done**

**Important**: Regular users will now get "Permission denied" if they try to open the sheet directly. This is correct! They should use the web app URL.

---

## Step 6: Share Web App URL with Users

Users access the app via the web app URL:

```
https://script.google.com/a/macros/threecolts.com/s/[YOUR_ID]/exec
```

They will:
- Sign in with their Google Workspace account
- See the Time Tracker interface
- Clock in/out, take breaks, track time
- All actions write to the sheet using YOUR permissions

---

## Testing

### Test as Admin (You)

1. Open the web app URL in your browser
2. Verify you can:
   - ✅ Clock in/out
   - ✅ See admin panel
   - ✅ View your profile
3. Open the sheet directly and verify the time entry was recorded

### Test as Regular User

1. Have a regular user open the web app URL (or use incognito mode)
2. They should be able to:
   - ✅ Sign in with their Workspace account
   - ✅ Clock in/out
   - ✅ See the time tracker interface
   - ❌ Admin panel should be hidden
3. They should NOT be able to:
   - ❌ Open the Google Sheet directly (permission denied)
   - ❌ Edit the sheet
4. Verify their time entry appears in the sheet (check as admin)

---

## What Changed?

### Code Changes
- `Session.getActiveUser()` → `Session.getEffectiveUser()` in 7 functions
- `getTodaysCalendarEvents()` now returns empty array (calendar disabled)

### Deployment Changes
- **Execute as: Me** - Script runs with owner's permissions
- Users don't need sheet access
- Calendar feature disabled (no Admin Console access needed)

---

## Troubleshooting

### "Permission denied" when users try to access

**Cause**: Deployment settings incorrect or users not in your organization

**Solution**:
- Verify "Execute as: Me" is selected
- Check "Who has access" includes your organization
- Make sure users are signed in with correct Workspace account

### Sheet writes fail / "You do not have permission to edit"

**Cause**: Script not running with your permissions

**Solution**:
- MUST deploy as "Execute as: Me"
- Verify YOU (the owner) have edit access to the sheet
- Redeploy with correct settings

### Users see authorization dialog every time

**Cause**: Cookies blocked or private browsing mode

**Solution**:
- Users should allow cookies from script.google.com
- Don't use incognito mode for regular use
- After first authorization, it should remember

### "This app isn't verified" warning

**Cause**: Apps Script projects created by personal accounts show this

**Solution**:
- This is normal and safe for your own scripts
- Click **Advanced** → **Go to [Project name] (unsafe)**
- Only show this to trusted users in your organization

---

## Security Notes

### What Users CAN Do:
✅ Use the web app to clock in/out
✅ Track their time through the app interface
✅ Upload their own profile pictures/backgrounds
✅ Delete their own backgrounds

### What Users CANNOT Do:
❌ Open the Google Sheet directly
❌ Edit the sheet or modify formulas
❌ See other users' data in the sheet
❌ Access admin features

### What Admins CAN Do:
✅ Everything users can do
✅ Open and edit the sheet directly
✅ Generate new week sheets
✅ Access admin panel
✅ View all users' status

---

## Re-enabling Calendar (Future)

If you gain access to Google Workspace Admin Console in the future:

1. Follow **SETUP_DOMAIN_WIDE_DELEGATION.md**
2. In `code.gs`, uncomment the calendar implementation in `getTodaysCalendarEvents()`
3. Remove the `return [];` line
4. Redeploy the app

---

## Summary Checklist

Before sharing with users:

- [ ] Code has latest changes (Session.getEffectiveUser)
- [ ] Calendar function returns empty array
- [ ] Deployed as "Execute as: Me"
- [ ] "Who has access" set correctly for your organization
- [ ] Web app URL copied
- [ ] WEB_APP_URL in code.gs updated (if changed)
- [ ] Sheet permissions restricted (only admins have direct access)
- [ ] Tested as admin - works correctly
- [ ] Tested as regular user - can clock in/out without sheet access
- [ ] Regular user CANNOT open sheet directly

---

## What's Next?

Your Time Tracker is ready to use!

**Key Points:**
- ✅ Main problem solved: Users don't need sheet edit permissions
- ✅ Simple deployment: No Admin Console access needed
- ✅ Secure: Sheet structure protected
- ✅ Works immediately: No complex setup required

**Trade-off:**
- ❌ No calendar integration (can be added later with admin access)

---

**Deployment Mode**: Simple (No Calendar)
**Last Updated**: 2025-11-09
**Status**: Ready to deploy ✅
