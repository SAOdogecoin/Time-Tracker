# Time Tracker - Permission Solutions

## The Problem Explained

There's a fundamental conflict in Google Apps Script web apps:

| Deployment Mode | Sheet Write Access | Calendar Access | User Needs Sheet Permission? |
|----------------|-------------------|-----------------|------------------------------|
| **Execute as: Me** | ✅ Owner's permissions | ❌ Owner's calendar only | **NO** |
| **Execute as: User** | ❌ User's permissions | ✅ User's own calendar | **YES** |

### Why This Happens

- **Sheet writes**: Require edit permissions. When running "Execute as: Me", the script uses the owner's (admin's) permissions.
- **Calendar access**: `Calendar.Events.list(userEmail)` only works for:
  - The executing user's own calendar, OR
  - Calendars explicitly shared with the executing user, OR
  - When domain-wide delegation is enabled (Workspace admin feature)

## Solution Options

### ✅ Option 1: Service Account with Domain-Wide Delegation (RECOMMENDED)

**Pros**: Solves both problems completely
**Cons**: Requires Google Workspace Admin privileges
**Complexity**: Medium

#### How It Works
1. Deploy with "Execute as: Me"
2. Enable domain-wide delegation for your Apps Script project
3. Script can access any user's calendar while running with owner's sheet permissions

#### Setup Steps

**Step 1: Enable Google Cloud Project**
1. Open Apps Script editor
2. Go to **Project Settings** (gear icon)
3. Under "Google Cloud Platform (GCP) Project", click **Change project**
4. Note your project number or create a new project

**Step 2: Enable Calendar API**
1. In Apps Script, go to **Services** (+ icon)
2. Add **Google Calendar API**

**Step 3: Configure OAuth Consent** (if not already done)
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Select your project
3. Navigate to **APIs & Services** → **OAuth consent screen**
4. Select **Internal** (for Workspace users only)
5. Fill in application name: "Time Tracker"
6. Add scopes:
   - `https://www.googleapis.com/auth/calendar.readonly`
   - `https://www.googleapis.com/auth/spreadsheets`
   - `https://www.googleapis.com/auth/drive.file`

**Step 4: Create Service Account Credentials**
1. In Google Cloud Console, go to **IAM & Admin** → **Service Accounts**
2. Click **Create Service Account**
3. Name: "time-tracker-service"
4. Click **Create and Continue**
5. Grant role: **Service Account User**
6. Click **Done**
7. Click on the service account email
8. Note the **Unique ID** (numeric ID)

**Step 5: Enable Domain-Wide Delegation** (Requires Workspace Admin)
1. Still in the service account details, click **Advanced settings**
2. Check **Enable Google Workspace Domain-wide Delegation**
3. Copy the **Client ID**

**Step 6: Authorize in Workspace Admin Console** (Workspace Admin Required)
1. Go to [Google Admin Console](https://admin.google.com/)
2. Navigate to **Security** → **Access and data control** → **API controls**
3. Click **Manage Domain Wide Delegation**
4. Click **Add new**
5. Paste the **Client ID** from Step 5
6. Add OAuth scopes:
   ```
   https://www.googleapis.com/auth/calendar.readonly,
   https://www.googleapis.com/auth/spreadsheets,
   https://www.googleapis.com/auth/drive.file
   ```
7. Click **Authorize**

**Step 7: Deploy Web App**
1. In Apps Script, click **Deploy** → **New deployment**
2. Select **Web app**
3. **Execute as**: **Me** (owner)
4. **Who has access**: **Anyone within [your organization]**
5. Click **Deploy**

**Result**:
- ✅ Users can use the app without sheet edit permissions
- ✅ Users see their own Google Calendar events
- ✅ All time entries are written to the sheet with owner permissions

---

### ⚠️ Option 2: Separate Service for Sheet Writes (Complex)

**Pros**: Works without admin access
**Cons**: Complex implementation, requires maintaining two deployments
**Complexity**: High

#### How It Works
1. Main web app: Deploy "Execute as: User" for calendar access
2. Separate script: Deploy "Execute as: Me" as an API endpoint for sheet writes
3. Main app calls the API endpoint to write data

#### Implementation Steps

This requires significant code changes:

**A. Create a separate Apps Script project for writes**
1. Create new Apps Script project
2. Deploy it as an API endpoint with "Execute as: Me"
3. Secure it with a secret key or authentication token

**B. Modify main app to call the API**
- Frontend calls main app for calendar data (runs as user)
- Frontend calls write API for time entries (runs as owner)

This is complex and not recommended unless absolutely necessary.

---

### ⚠️ Option 3: Remove Calendar Integration (Simple)

**Pros**: Simple, solves the sheet permission problem immediately
**Cons**: Loses calendar feature
**Complexity**: Low

#### How It Works
1. Comment out or remove calendar-related code
2. Deploy with "Execute as: Me"
3. Users can track time without seeing calendar events

#### Code Changes Required

```javascript
// In code.gs, comment out the calendar function:
/*
function getTodaysCalendarEvents() {
  // ... entire function commented out
}
*/

// Return empty array instead:
function getTodaysCalendarEvents() {
  return [];
}
```

**Result**:
- ✅ Users can use the app without sheet edit permissions
- ❌ No calendar integration

---

### ⚠️ Option 4: Hybrid Approach - Optional Calendar Access

**Pros**: Users can optionally grant calendar access
**Cons**: Calendar feature becomes opt-in, requires user permission dialog
**Complexity**: Medium

#### How It Works
1. Deploy with "Execute as: Me" for sheet writes
2. Use client-side Google Calendar API for calendar access
3. Users see a one-time permission dialog to allow calendar access

#### Implementation Changes Required

1. Remove server-side calendar calls
2. Add client-side Google Calendar API calls using `gapi.client`
3. Request calendar permissions in frontend

This requires significant frontend code changes.

---

### ⚠️ Option 5: Keep Current Approach, Grant Users "Commenter" Access

**Pros**: Simplest, no code changes needed
**Cons**: Doesn't fully solve the permission problem
**Complexity**: Very Low

#### How It Works
1. Keep "Execute as: User accessing the web app"
2. Grant all users **Commenter** access to the sheet (not Editor)
3. Users can see the sheet but can't directly edit it
4. The app writes using their commenter permissions...

**Wait - this won't work!** Commenters can't write to sheets via Apps Script either.

---

## Recommendation

Based on your requirements:

### If you have Workspace Admin access:
→ **Use Option 1 (Domain-Wide Delegation)** ← BEST SOLUTION

### If you DON'T have Workspace Admin access:

**Ask yourself**: Is the calendar feature essential?

- **Calendar is essential** → Use Option 4 (Client-side calendar with permission dialog)
- **Calendar is nice-to-have** → Use Option 3 (Remove calendar feature)
- **You have dev resources** → Use Option 2 (Separate write service)

## Next Steps

Please let me know:
1. **Do you have Google Workspace Admin access?** (or can you request it?)
2. **Is the calendar integration critical?** (or can it be removed/optional?)
3. **What's your preference?** (Which option above fits your needs best?)

Once you decide, I'll implement the chosen solution.
