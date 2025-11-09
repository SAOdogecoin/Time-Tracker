# Shared Calendar Setup Guide

## Overview

This guide explains how to enable the **Shared Calendar** feature in your Time Tracker app. This allows you to show calendar events from specific users (like admins or managers) to all app users.

**Perfect for showing:**
- Team-wide meetings
- Office events (all-hands, holidays)
- Manager/admin availability
- Important company events

**No domain-wide delegation needed!** ✅

---

## How It Works

1. **Specific users share their calendars** with you (the script owner)
2. **You add their emails to the code** (SHARED_CALENDAR_EMAILS array)
3. **App shows those calendars' events to everyone** (all users see the same events)

### Important Notes

- ❌ Users will NOT see their own personal calendar events
- ✅ Users WILL see events from the shared calendars you configure
- ✅ All users see the same calendar events
- ✅ No Admin Console access needed

---

## Setup Steps

### Step 1: Ask Users to Share Their Calendars

Have the specific users (e.g., admins, managers) share their Google Calendar with you:

**Instructions for them:**
1. Open [Google Calendar](https://calendar.google.com)
2. On the left, find "My calendars"
3. Hover over their calendar name and click the three dots (⋮)
4. Click **Settings and sharing**
5. Scroll down to "Share with specific people or groups"
6. Click **+ Add people and groups**
7. Enter your email (script owner): `cantiporta@threecolts.com`
8. Set permission: **See all event details**
9. Click **Send**

**You'll receive an email** with the calendar share invitation. Accept it.

### Step 2: Update SHARED_CALENDAR_EMAILS in Code

1. Open your Time Tracker Google Sheet
2. Click **Extensions** → **Apps Script**
3. In `code.gs`, find lines 16-19:

   ```javascript
   const SHARED_CALENDAR_EMAILS = [
     // 'cantiporta@threecolts.com',  // Uncomment to show this user's calendar
     // 'apelagio@threecolts.com',     // Uncomment to show this user's calendar
   ];
   ```

4. **Uncomment and add the emails** of users who shared their calendars:

   ```javascript
   const SHARED_CALENDAR_EMAILS = [
     'cantiporta@threecolts.com',  // Show this user's calendar
     'apelagio@threecolts.com',     // Show this user's calendar
   ];
   ```

5. Click **Save** (Ctrl+S or Cmd+S)

### Step 3: Verify Deployment Settings

Make sure your app is deployed correctly:

1. In Apps Script, click **Deploy** → **Manage deployments**
2. Verify these settings:
   - **Execute as**: **Me (cantiporta@threecolts.com)** ← MUST BE "Me"
   - **Who has access**: **Anyone within threecolts.com**

3. If you need to change settings:
   - Click **Edit** (pencil icon)
   - Update the settings
   - Select **New version**
   - Click **Deploy**

### Step 4: Test the Calendar Feature

1. Open the web app URL in your browser
2. You should see calendar events from the shared calendars
3. Have another user test it - they should see the same events
4. If no events show up, check:
   - Calendar sharing is set up correctly
   - The calendars have events for today
   - Check Apps Script logs: **Executions** tab

---

## Configuration Examples

### Example 1: Show Only Your Calendar

```javascript
const SHARED_CALENDAR_EMAILS = [
  'cantiporta@threecolts.com',  // Show only super admin's calendar
];
```

### Example 2: Show Multiple Admin Calendars

```javascript
const SHARED_CALENDAR_EMAILS = [
  'cantiporta@threecolts.com',  // Super admin calendar
  'apelagio@threecolts.com',     // Admin calendar
];
```

### Example 3: Show Team Calendar + Manager

```javascript
const SHARED_CALENDAR_EMAILS = [
  'team@threecolts.com',         // Shared team calendar account
  'manager@threecolts.com',      // Manager's calendar
];
```

### Example 4: Disable Calendar Completely

```javascript
const SHARED_CALENDAR_EMAILS = [
  // Leave empty to disable calendar feature
];
```

---

## Updating Calendar Configuration

To add or remove calendars:

1. Update `SHARED_CALENDAR_EMAILS` array in `code.gs`
2. Click **Save**
3. **Redeploy**:
   - **Deploy** → **Manage deployments**
   - Click **Edit**
   - Select **New version**
   - Click **Deploy**
4. Users may need to refresh the app to see changes

---

## Troubleshooting

### No Events Showing Up

**Check 1: Calendar Sharing**
- Verify the user shared their calendar with you (script owner)
- Check your Google Calendar - can you see their events?
- Make sure permission is "See all event details" (not just "See only free/busy")

**Check 2: Email in Array**
- Verify the email in SHARED_CALENDAR_EMAILS is correct
- No typos, correct domain
- Remove any extra spaces

**Check 3: Events Exist**
- Check if there are actually events today in those calendars
- Try adding a test event to verify

**Check 4: Deployment**
- Must be deployed as "Execute as: Me"
- Make sure you redeployed after changing SHARED_CALENDAR_EMAILS

**Check 5: Calendar API**
- Go to Apps Script → **Services**
- Make sure "Google Calendar API" is added
- If not, click **+ Add a service** → **Google Calendar API**

### Error: "Calendar not found"

**Cause**: The calendar email doesn't exist or wasn't shared with you

**Solution**:
- Double-check the email address
- Ask the user to share again
- Accept the sharing invitation in your email

### Events Show but Google Meet Links Missing

**Cause**: Some events don't have Google Meet links

**Solution**:
- This is normal - only events with Meet links will show the link
- Events without Meet links will show `gmeetLink: null`

### Users See Different Events

**Cause**: This shouldn't happen - all users should see the same events

**Solution**:
- Check that you're using SHARED_CALENDAR_EMAILS (not per-user logic)
- Verify deployment is "Execute as: Me"
- Have users refresh the app (hard refresh: Ctrl+Shift+R)

---

## Calendar Event Format

Events returned to the frontend include:

```javascript
{
  id: "event123",
  title: "Team Meeting",
  startTime: "2025-11-09T14:00:00-08:00",  // ISO format
  endTime: "2025-11-09T15:00:00-08:00",    // ISO format
  gmeetLink: "https://meet.google.com/abc-defg-hij",  // or null
  calendarOwner: "cantiporta@threecolts.com"  // Who owns this calendar
}
```

The `calendarOwner` field helps users know whose calendar the event is from.

---

## Privacy & Security Notes

### What Users Can See:
✅ Event titles from shared calendars
✅ Event times (start/end)
✅ Google Meet links (if event has one)
✅ Which calendar the event is from (calendarOwner)

### What Users CANNOT See:
❌ Event descriptions (not included)
❌ Event attendees (not included)
❌ Private event details
❌ Their own personal calendar events

### Admin Control:
- Only calendars YOU choose are shown
- You control which users' calendars to share
- You can remove calendars anytime by updating the array

---

## Comparison: Shared vs Personal Calendars

| Feature | Shared Calendars (This Guide) | Personal Calendars (Requires Admin) |
|---------|------------------------------|-------------------------------------|
| **Setup Difficulty** | Easy (5 minutes) | Hard (requires Admin Console) |
| **Admin Access Needed** | ❌ No | ✅ Yes (domain-wide delegation) |
| **What Users See** | Same events for everyone | Each user sees their own calendar |
| **Use Case** | Team events, office calendar | Personal time management |
| **Calendar Sharing Required** | ✅ Yes (specific users share with you) | ❌ No (automatic via delegation) |

---

## Example Use Cases

### Use Case 1: Office-Wide Events
- Create a shared calendar account: `events@threecolts.com`
- Add all company events, holidays, all-hands meetings
- Share with script owner
- Add to SHARED_CALENDAR_EMAILS
- All employees see important events in the Time Tracker

### Use Case 2: Manager Availability
- Manager shares their calendar with you
- Add manager's email to SHARED_CALENDAR_EMAILS
- Team can see when manager has meetings
- Helps with scheduling and availability awareness

### Use Case 3: Multiple Department Calendars
- Sales dept has `sales@threecolts.com` calendar
- Engineering has `eng@threecolts.com` calendar
- Add both to SHARED_CALENDAR_EMAILS
- All users see cross-functional meetings

---

## Summary Checklist

Before enabling shared calendars:

- [ ] Specific users have shared their calendars with you (script owner)
- [ ] You can see their events in your Google Calendar
- [ ] Added their emails to SHARED_CALENDAR_EMAILS array
- [ ] Saved code.gs
- [ ] Redeployed the app with new version
- [ ] Verified deployment is "Execute as: Me"
- [ ] Google Calendar API is enabled in Services
- [ ] Tested - events show up in the app

---

## Disabling Calendar Feature

To disable completely:

```javascript
const SHARED_CALENDAR_EMAILS = [
  // Leave empty - no calendars configured
];
```

The app will work perfectly without any calendar events.

---

**Feature**: Shared Calendars (No Admin Access Required)
**Last Updated**: 2025-11-09
**Status**: Ready to use ✅
