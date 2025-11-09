# Event Creator Filter Setup Guide

## Overview

This guide explains how to set up **calendar event filtering by creator** in your Time Tracker app. This allows you to show only events created by specific users (like Gemma Bianchi) to all app users, while keeping other events visible only to admins.

**Perfect Privacy Feature!** ✅

---

## How It Works

The app reads calendar events and checks **who CREATED each event** (the organizer):

- ✅ **Event created by Gemma Bianchi** → ALL users see it
- ✅ **Event created by you** → ALL users see it (if you're in the list)
- ❌ **Event created by someone else** → Only ADMINS see it

### Example Scenario

**Your calendar has these events today:**
1. "Team Standup" - created by Gemma Bianchi (gbnchi@threecolts.com)
2. "Client Meeting" - created by John Doe (john@client.com)
3. "1-on-1 with Manager" - created by your manager

**What regular users see:**
- ✅ "Team Standup" (Gemma created it - she's in ALLOWED_EVENT_CREATORS)

**What admins see:**
- ✅ "Team Standup" (Gemma created it)
- ✅ "Client Meeting" (admin can see all events)
- ✅ "1-on-1 with Manager" (admin can see all events)

---

## Benefits

✅ **Privacy-focused**: Users only see events from approved organizers
✅ **Flexible**: Show company-wide meetings without exposing personal events
✅ **No Admin Console needed**: Works without domain-wide delegation
✅ **Simple setup**: Just add emails to a list
✅ **Admin override**: Admins always see all events

---

## Setup Steps

### Step 1: Have Users Share Their Calendars with You

First, users whose events you want to check must share their calendars with you (the script owner).

**Common approach**: Use your own calendar (already have access)

**For other calendars**: Ask users to share their calendar with you:

1. They open [Google Calendar](https://calendar.google.com)
2. Settings → Share with specific people → Add your email
3. Permission: "See all event details"
4. Click Send

**You'll receive an invitation** - accept it to gain access.

### Step 2: Configure CALENDARS_TO_CHECK

Open `code.gs` and find lines 10-14:

```javascript
const CALENDARS_TO_CHECK = [
  'primary',  // Your own calendar (the script owner's calendar)
  // 'cantiporta@threecolts.com',  // Add more calendars if they're shared with you
  // 'apelagio@threecolts.com',
];
```

**Options:**
- `'primary'` - Your own calendar (already enabled by default)
- Add specific emails of users who shared their calendars with you

**Example - Check your calendar only:**
```javascript
const CALENDARS_TO_CHECK = [
  'primary',  // Just check your calendar
];
```

**Example - Check multiple calendars:**
```javascript
const CALENDARS_TO_CHECK = [
  'primary',  // Your calendar
  'team@threecolts.com',  // Shared team calendar
];
```

### Step 3: Configure ALLOWED_EVENT_CREATORS

This is the KEY setting! Find lines 27-30:

```javascript
const ALLOWED_EVENT_CREATORS = [
  // 'gbnchi@threecolts.com',  // Gemma Bianchi - uncomment to show her events to everyone
  // 'cantiporta@threecolts.com',  // Add more allowed creators
];
```

**Uncomment and add emails** of users whose events should be visible to ALL users:

```javascript
const ALLOWED_EVENT_CREATORS = [
  'gbnchi@threecolts.com',  // Gemma Bianchi's events → everyone sees
  'cantiporta@threecolts.com',  // Your events → everyone sees
];
```

### Step 4: Save and Deploy

1. Click **Save** (Ctrl+S or Cmd+S) in Apps Script
2. Click **Deploy** → **Manage deployments**
3. Click **Edit** (pencil icon)
4. Select **New version**
5. Click **Deploy**
6. Make sure deployment is:
   - **Execute as**: **Me** (your email)
   - **Who has access**: **Anyone within threecolts.com**

---

## Configuration Examples

### Example 1: Show Only Gemma Bianchi's Events

```javascript
const CALENDARS_TO_CHECK = [
  'primary',  // Check your calendar
];

const ALLOWED_EVENT_CREATORS = [
  'gbnchi@threecolts.com',  // Only Gemma's events show to everyone
];
```

**Result**: If Gemma creates a meeting on your calendar, all users see it. Other events? Only admins.

### Example 2: Show Events from Multiple Team Leads

```javascript
const CALENDARS_TO_CHECK = [
  'primary',  // Your calendar
  'team@threecolts.com',  // Shared team calendar
];

const ALLOWED_EVENT_CREATORS = [
  'gbnchi@threecolts.com',  // Gemma's events
  'cantiporta@threecolts.com',  // Your events
  'apelagio@threecolts.com',  // Admin's events
];
```

**Result**: Events created by Gemma, you, or apelagio are visible to everyone. Others? Only admins.

### Example 3: Only Admins See All Events (No Public Events)

```javascript
const CALENDARS_TO_CHECK = [
  'primary',
];

const ALLOWED_EVENT_CREATORS = [
  // Empty - no events shown to regular users
];
```

**Result**: Calendar disabled for regular users. Admins see all events from your calendar.

---

## Understanding the Filtering

### The Organizer Field

Every Google Calendar event has an **organizer** - the person who created the event.

```javascript
// Example event data from Google Calendar API:
{
  "summary": "Team Meeting",
  "organizer": {
    "email": "gbnchi@threecolts.com",  // Who created it
    "displayName": "Gemma Bianchi"
  },
  "attendees": [
    { "email": "user1@threecolts.com" },
    { "email": "user2@threecolts.com" }
  ]
}
```

**The script checks**: Is `organizer.email` in `ALLOWED_EVENT_CREATORS`?
- **YES** → Show to everyone
- **NO** → Show only to admins

### Special Cases

**Case 1: You're invited to someone else's meeting**
- Organizer: someone@external.com
- Your calendar shows this event (you're an attendee)
- Script sees organizer is NOT in ALLOWED_EVENT_CREATORS
- Result: Only admins see it ✅

**Case 2: Gemma creates a meeting and invites you**
- Organizer: gbnchi@threecolts.com
- Your calendar shows this event
- Script sees organizer IS in ALLOWED_EVENT_CREATORS
- Result: Everyone sees it ✅

**Case 3: External meeting on a shared team calendar**
- Organizer: client@external.com
- Team calendar shows this event
- Script sees organizer is NOT in ALLOWED_EVENT_CREATORS
- Result: Only admins see it ✅

---

## Testing

### Test as Regular User

1. Have a regular user (non-admin) open the web app
2. They should only see events created by users in ALLOWED_EVENT_CREATORS
3. Other events should be hidden

### Test as Admin

1. Open the web app as an admin (you or apelagio)
2. You should see ALL events from your calendar
3. Including events not created by allowed users

### Create Test Events

**To test the filter:**

1. Create a test event on your calendar as yourself
2. Add it to ALLOWED_EVENT_CREATORS (your email)
3. Regular users should see it ✅

4. Have someone else create an event and invite you
5. Don't add their email to ALLOWED_EVENT_CREATORS
6. Regular users should NOT see it ❌
7. Admins should see it ✅

---

## Troubleshooting

### No Events Showing Up for Anyone

**Check 1: ALLOWED_EVENT_CREATORS not empty**
```javascript
const ALLOWED_EVENT_CREATORS = [
  'gbnchi@threecolts.com',  // Must have at least one
];
```

**Check 2: CALENDARS_TO_CHECK not empty**
```javascript
const CALENDARS_TO_CHECK = [
  'primary',  // Must have at least one
];
```

**Check 3: Events exist today**
- Check your Google Calendar
- Make sure there are events today
- Make sure they're timed events (not all-day)

**Check 4: Correct organizer email**
- Open the event in Google Calendar
- Check who created it (organizer)
- Make sure that email is in ALLOWED_EVENT_CREATORS

### Regular Users See Events They Shouldn't

**Cause**: The event organizer is in ALLOWED_EVENT_CREATORS

**Check**:
- Open the event details in Google Calendar
- Look at the organizer email
- Is it in your ALLOWED_EVENT_CREATORS list?
- If yes, that's expected behavior - remove it if you don't want to show their events

### Admins Don't See All Events

**Cause 1**: Admin email not in ADMIN_EMAILS or SUPER_ADMIN_EMAILS

**Solution**: Check lines 1-2 of code.gs:
```javascript
const SUPER_ADMIN_EMAILS = ['cantiporta@threecolts.com'];
const ADMIN_EMAILS = ['apelagio@threecolts.com'];
```

**Cause 2**: Calendar not shared or accessible

**Solution**: Make sure the calendar in CALENDARS_TO_CHECK is accessible to you (the script owner)

### Error: "Calendar not found"

**Cause**: Calendar ID in CALENDARS_TO_CHECK doesn't exist or isn't shared with you

**Solution**:
- Check the email/calendar ID is correct
- If it's someone else's calendar, make sure they shared it with you
- Accept the sharing invitation in your email

---

## Event Data Returned to Frontend

Events include:

```javascript
{
  id: "event123",
  title: "Team Meeting",
  startTime: "2025-11-09T14:00:00-08:00",  // ISO format
  endTime: "2025-11-09T15:00:00-08:00",
  gmeetLink: "https://meet.google.com/abc-defg-hij",  // or null
  createdBy: "gbnchi@threecolts.com"  // Who created/organized this event
}
```

The `createdBy` field shows the organizer's email.

---

## Privacy Considerations

### What Regular Users See:
✅ Event titles from allowed creators
✅ Event times
✅ Google Meet links
✅ Who created the event (createdBy field)

### What Regular Users DON'T See:
❌ Events created by non-allowed users
❌ Event descriptions (not included in API response)
❌ Full attendee lists (not included in API response)
❌ Personal/private events from other people

### What Admins See:
✅ ALL events from checked calendars
✅ Regardless of creator
✅ Useful for monitoring and troubleshooting

---

## Use Cases

### Use Case 1: Show Team Leads' Meetings Only
- Add team leads to ALLOWED_EVENT_CREATORS
- Everyone sees when team leads schedule meetings
- Personal meetings and external meetings stay private

### Use Case 2: Company-Wide Events
- Designate one person (e.g., HR) to create company events
- Add their email to ALLOWED_EVENT_CREATORS
- All-hands, holidays, company events visible to everyone
- Other calendar events stay private

### Use Case 3: Department Calendar
- Create a shared calendar: engineering@threecolts.com
- Add to CALENDARS_TO_CHECK
- Add engineering leads to ALLOWED_EVENT_CREATORS
- Team sees important meetings, private ones stay hidden

---

## Comparison: Creator Filter vs Shared Calendars

| Feature | Creator Filter (This Guide) | Shared Calendars (Previous) |
|---------|----------------------------|------------------------------|
| **Privacy** | ✅ Excellent (filter by creator) | ⚠️ Shows all calendar events |
| **Flexibility** | ✅ Very flexible | ⚠️ All or nothing |
| **Use Case** | Mixed calendar (work + personal) | Dedicated team calendars only |
| **Setup** | Simple (2 arrays) | Simple (1 array) |
| **Admin Override** | ✅ Yes | ❌ No |

---

## Disabling Calendar Feature

To disable completely:

```javascript
const ALLOWED_EVENT_CREATORS = [
  // Leave empty - calendar disabled
];
```

The app will work perfectly without calendar events.

---

## Summary Checklist

Before going live:

- [ ] Decided which calendars to check (added to CALENDARS_TO_CHECK)
- [ ] Calendars are accessible to you (script owner)
- [ ] Added allowed event creators' emails to ALLOWED_EVENT_CREATORS
- [ ] Verified admin emails in SUPER_ADMIN_EMAILS and ADMIN_EMAILS
- [ ] Saved code.gs
- [ ] Redeployed with "Execute as: Me"
- [ ] Tested as regular user - only see allowed creators' events
- [ ] Tested as admin - see all events
- [ ] Created test events to verify filtering works

---

**Feature**: Event Creator Filtering (Privacy-Focused)
**Setup Time**: 5 minutes
**Requires**: No Admin Console access needed
**Status**: Ready to use ✅
