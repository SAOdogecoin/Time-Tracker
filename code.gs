const SUPER_ADMIN_EMAILS = ['cantiporta@threecolts.com'];
const ADMIN_EMAILS = ['apelagio@threecolts.com'];

function isSuperAdmin_() {
  try {
    const currentUserEmail = Session.getActiveUser().getEmail();
    return SUPER_ADMIN_EMAILS.includes(currentUserEmail);
  } catch (e) {
    return false;
  }
}

function isAdmin_() {
  try {
    const currentUserEmail = Session.getActiveUser().getEmail();
    return SUPER_ADMIN_EMAILS.includes(currentUserEmail) || ADMIN_EMAILS.includes(currentUserEmail);
  } catch (e) {
    return false;
  }
}

const TEMPLATE_SHEET_NAME = "Template";
const SCRIPT_TIMEZONE = Session.getScriptTimeZone();
const ASSET_FOLDER_NAME = "Time Tracker Assets";

const PREF_KEYS = {
    CLOCK_IN_REMINDER_TIME: 'clockInReminderTime',
    SKIP_WEEKEND_AUTOMATION: 'skipWeekendAutomation'
};

const LOG_KEYS = {
    AUTOMATION_LOG: 'automationLog',
    RECENT_AUTOMATION_EVENTS: 'recentAutomationEvents'
};
const MAX_LOG_ENTRIES = 100;

const DATE_NAME_MAPPING = [ { dateRow: 1, nameStartRow: 3, nameEndRow: 27 }, { dateRow: 28, nameStartRow: 30, nameEndRow: 54 }, { dateRow: 55, nameStartRow: 57, nameEndRow: 81 }, { dateRow: 82, nameStartRow: 84, nameEndRow: 108 }, { dateRow: 109, nameStartRow: 111, nameEndRow: 135 }, { dateRow: 136, nameStartRow: 138, nameEndRow: 144 }, { dateRow: 145, nameStartRow: 147, nameEndRow: 153 } ];
const COLUMN_MAP = { 'clockIn': { index: 2, name: 'Clock In' }, 'breakOut': { index: 3, name: 'Break Out' }, 'breakIn': { index: 4, name: 'Break In' }, 'lunchOut': { index: 5, name: 'Lunch Out' }, 'lunchIn': { index: 6, name: 'Lunch In' }, 'clockOut': { index: 7, name: 'Clock Out' }, 'hours': { index: 8, name: 'Hours' } };
const TIME_CELL_FORMAT = 'hh:mm:ss AM/PM';


function doGet(e) {
  try {
    let template = HtmlService.createTemplateFromFile('Index');
    template.isAdmin = isAdmin_();
    template.isSuperAdmin = isSuperAdmin_();
    
    const scriptProps = PropertiesService.getScriptProperties();
    template.savedBackground = scriptProps.getProperty('backgroundSelection') || 'default';
    template.userNameFromUrl = e.parameter.user || '';

    let htmlOutput = template.evaluate().addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).setTitle(`Time Tracker`);
    return htmlOutput;
  } catch (err) {
    Logger.log(`Error serving HTML: ${err}`);
    return HtmlService.createHtmlOutput(`<b>Error loading application:</b> ${err.message}.`);
  }
}


function getUserProfileData(name) {
  try {
    const preferences = getUserPreferences_(name);
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const yesterday = new Date(today);
    yesterday.setDate(yesterday.getDate() - 1);

    let activeRow = null;
    let activeSheet = null;

    const findActiveSession = (date) => {
      const sheet = getSheetForDate_(date);
      if (!sheet) return null;
      
      const userEntry = findNameDataForDate_(sheet, date)?.find(u => u.name === name);
      if (userEntry?.row) {
        const values = sheet.getRange(userEntry.row, 2, 1, 6).getValues()[0];
        const clockInIsDate = values[0] instanceof Date;
        const clockOutIsDate = values[5] instanceof Date;

        if (clockInIsDate && !clockOutIsDate) {
          return { row: userEntry.row, sheet: sheet };
        }
      }
      return null;
    };

    const activeSession = findActiveSession(today) || findActiveSession(yesterday);
    
    if (activeSession) {
      activeRow = activeSession.row;
      activeSheet = activeSession.sheet;
    }

    let timeEntries = {};
    let status = 'Offline';
    let totalSecondsToday = 0;
    let activeLogKey = null;
    let lastActionTimestamp = null;
    
    const contextSheet = activeSheet || getSheetForDate_(today);
    const contextRow = activeRow || (contextSheet ? findTodaysRowForUser_(name, contextSheet) : null);

    if (contextSheet && contextRow) {
      timeEntries = getTimeEntriesForUser_(contextSheet, contextRow) || {};
      status = determineStatusFromTimes(timeEntries);
      activeLogKey = determineActiveLogKeyFromTimes(timeEntries);
      lastActionTimestamp = getLastActionTimestamp_(contextSheet, contextRow, activeLogKey);
      
      // Calculate total seconds for today
      if (activeSheet && getSheetForDate_(today) && activeSheet.getSheetId() === getSheetForDate_(today).getSheetId()) {
         // Active session is on today's sheet - calculate ongoing time
         totalSecondsToday = calculateCurrentTotalHoursSeconds_(contextSheet, contextRow);
      } else if (status === 'Offline') {
         // User is offline - get completed hours from today's row
         const todaySheet = getSheetForDate_(today);
         const todayRow = todaySheet ? findTodaysRowForUser_(name, todaySheet) : null;
         if (todayRow) {
            const todayHours = todaySheet.getRange(todayRow, COLUMN_MAP.hours.index).getValue();
            if (typeof todayHours === 'number') {
               totalSecondsToday = todayHours * 3600;
            }
         }
      } else if (activeSheet) {
         // EDGE CASE: Active session started yesterday (e.g., clocked in at 11:50 PM, still working after midnight)
         // Calculate time worked but attribute to the day the session started (yesterday's sheet)
         // Display the ongoing time for status display purposes
         totalSecondsToday = calculateCurrentTotalHoursSeconds_(activeSheet, activeRow);
         Logger.log(`Midnight session edge case: ${name} has active session from previous day, showing ongoing time: ${(totalSecondsToday / 3600).toFixed(2)} hours`);
      }
    }

    const profileData = { 
      name: name, 
      preferences: preferences, 
      profilePicFileId: PropertiesService.getScriptProperties().getProperty(`profilePic_name_${sanitizeKey_(name)}`), 
      status: status, 
      timeEntries: timeEntries, 
      totalSecondsToday: totalSecondsToday, 
      actualDataRowForStatus: activeRow,
      todaysRow: contextRow,
      activeLogKey: activeLogKey,
      lastActionTimestamp: lastActionTimestamp
    };
    
    return { success: true, data: profileData };
  } catch (e) {
    Logger.log(`Error in getUserProfileData for ${name}: ${e.stack}`);
    return { success: false, error: e.toString() };
  }
}

function getSheetForDate_(date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  const dayNum = d.getDay() || 7;
  d.setDate(d.getDate() + 4 - dayNum);
  const yearStart = new Date(d.getFullYear(), 0, 1);
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  const targetSheetName = `Week${weekNo}`;
  return ss.getSheetByName(targetSheetName);
}

function findTodaysRowForUser_(userName, sheet) {
  const targetSheet = sheet || getOrCreateTargetSheet_();
  if (!targetSheet) return null;
  const today = new Date();
  today.setHours(0,0,0,0);
  const userToday = findNameDataForDate_(targetSheet, today)?.find(u => u.name === userName);
  return userToday?.row || null;
}

function determineActiveLogKeyFromTimes(times) {
  const has = (key) => times && times[key] && String(times[key]).trim() !== "" && String(times[key]).trim() !== "--";
  if (has('clockOut')) return 'clockOut';
  if (has('lunchIn')) return 'lunchIn';
  if (has('lunchOut')) return 'lunchOut';
  if (has('breakIn')) return 'breakIn';
  if (has('breakOut')) return 'breakOut';
  if (has('clockIn')) return 'clockIn';
  return null;
}

function getLastActionTimestamp_(sheet, row, activeLogKey) {
  if (!sheet || !row || !activeLogKey || !COLUMN_MAP[activeLogKey]) return null;
  try {
    const dateValue = sheet.getRange(row, COLUMN_MAP[activeLogKey].index).getValue();
    return (dateValue instanceof Date) ? dateValue.getTime() : null;
  } catch (e) {
    return null;
  }
}

// Hourly and daily checks have been removed as they were Telegram-dependent
// Users will receive in-app notifications instead when automation runs

function getInitialLoadData(userName) {
  const result = { success: false, nameList: [], initialUserData: null, dynamicGreeting: null, error: null };
  try {
    const allUsersResult = getAllUsersStatus();
    if (!allUsersResult.success) {
      throw new Error(allUsersResult.error);
    }
    result.nameList = allUsersResult.users.map(u => ({ name: u.name }));
    const initialUserName = userName || (result.nameList.length > 0 ? result.nameList[0].name : null);
    if (initialUserName) {
      const profileDataResult = getUserProfileData(initialUserName);
      if (profileDataResult.success) {
        result.initialUserData = profileDataResult.data;
      }
      result.dynamicGreeting = getDynamicGreeting(initialUserName);
    }
    result.success = true;
  } catch (e) {
    result.error = e.message;
  }
  return result;
}

function getAllUsersStatus() {
  const result = { success: false, users: [], error: null };
  try {
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const yesterday = new Date(today);
    yesterday.setDate(today.getDate() - 1);

    const sheetToday = getSheetForDate_(today);
    const sheetYesterday = getSheetForDate_(yesterday);

    const todaysUsers = sheetToday ? findNameDataForDate_(sheetToday, today) || [] : [];
    const yesterdaysUsers = sheetYesterday ? findNameDataForDate_(sheetYesterday, yesterday) || [] : [];
    
    const finalUserMap = {};

    const processUser = (user, sheet) => {
      const times = getTimeEntriesForUser_(sheet, user.row);
      const status = determineStatusFromTimes(times);
      let hoursWorked = 0;
      if (status !== 'Offline') {
        hoursWorked = calculateCurrentTotalHoursSeconds_(sheet, user.row) / 3600;
      } else if (times?.clockOut && times.clockOut !== '--') {
        const hoursValue = sheet.getRange(user.row, COLUMN_MAP.hours.index).getValue();
        if (typeof hoursValue === 'number') {
          hoursWorked = hoursValue;
        }
      }
      return { name: user.name, row: user.row, status: status, hoursWorked: hoursWorked, profilePicFileId: PropertiesService.getScriptProperties().getProperty(`profilePic_name_${sanitizeKey_(user.name)}`), actualDataRowForStatus: user.row };
    };

    if (sheetYesterday) {
      yesterdaysUsers.forEach(user => {
        const data = processUser(user, sheetYesterday);
        if (data.status !== 'Offline') {
          finalUserMap[user.name] = data;
        }
      });
    }
    if(sheetToday){
      todaysUsers.forEach(user => {
        if (!finalUserMap[user.name]) {
          finalUserMap[user.name] = processUser(user, sheetToday);
        }
      });
    }

    result.users = Object.values(finalUserMap).sort((a,b) => a.name.localeCompare(b.name));
    result.success = true;
  } catch (e) {
    result.error = e.message;
  }
  return result;
}

function processTimeEntry(userName, actionKey, timestamp) {
  const scriptNow = timestamp || new Date();

  try {
    let sheet;
    let targetRowNumber;

    if (actionKey === 'clockIn') {
      sheet = getOrCreateTargetSheet_(); 
      if (!sheet) throw new Error('Could not access the timesheet for today.');
      targetRowNumber = findTodaysRowForUser_(userName, sheet);
    } else {
      const now = new Date();
      const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
      const yesterday = new Date(today);
      yesterday.setDate(yesterday.getDate() - 1);

      const findActiveSession = (date) => {
        const sessionSheet = getSheetForDate_(date);
        if (!sessionSheet) return null;
        const userEntry = findNameDataForDate_(sessionSheet, date)?.find(u => u.name === userName);
        if (userEntry?.row) {
          const values = sessionSheet.getRange(userEntry.row, 2, 1, 6).getValues()[0];
          if ((values[0] instanceof Date) && !(values[5] instanceof Date)) {
            return { row: userEntry.row, sheet: sessionSheet };
          }
        }
        return null;
      };
      
      const activeSession = findActiveSession(today) || findActiveSession(yesterday);

      if (activeSession) {
        sheet = activeSession.sheet;
        targetRowNumber = activeSession.row;
      }
    }

    if (!targetRowNumber || !sheet) {
      throw new Error(`Could not find an active session or today's schedule for ${userName}.`);
    }

    const column = COLUMN_MAP[actionKey];
    if (!column) throw new Error('Invalid action.');

    if (actionKey === 'clockIn') {
      const existingTimes = getTimeEntriesForUser_(sheet, targetRowNumber);
      if (existingTimes && existingTimes.clockOut && existingTimes.clockOut.trim() !== '--') {
        throw new Error('You have already clocked out for today. Cannot clock in again.');
      }
    }

    sheet.getRange(targetRowNumber, column.index).setValue(scriptNow).setNumberFormat(TIME_CELL_FORMAT);
    SpreadsheetApp.flush();
    
    if (actionKey === 'clockOut') {
      recalculateHours(targetRowNumber, sheet);
    }
    
    const updatedProfileResult = getUserProfileData(userName);
    if (!updatedProfileResult.success) {
      throw new Error('Failed to retrieve updated user profile after time entry.');
    }
    
    const dynamicGreeting = getDynamicGreeting(userName);

    return { success: true, updatedUserData: updatedProfileResult.data, dynamicGreeting: dynamicGreeting };
  } catch (error) {
    Logger.log(`Error in processTimeEntry for ${userName} (${actionKey}): ${error.stack}`);
    return { success: false, error: error.message };
  }
}

/**
 * Check if Calendar API is available and authorized
 * @returns {Object} Status object with success flag and message
 */
function checkCalendarAPIStatus() {
  try {
    // Try to access the Calendar service
    Calendar.Events.list('primary', {
      maxResults: 1,
      timeMin: new Date().toISOString()
    });
    return {
      success: true,
      available: true,
      message: 'Calendar API is available and authorized'
    };
  } catch (e) {
    const errorMsg = e.toString();

    // Check if it's a permission/authorization error
    if (errorMsg.includes('Authorization') || errorMsg.includes('permission')) {
      return {
        success: true,
        available: false,
        needsAuth: true,
        message: 'Calendar API needs authorization. Click "Enable Calendar" to grant access.'
      };
    }

    // Check if the Calendar service is not enabled
    if (errorMsg.includes('Calendar') || errorMsg.includes('not found') || errorMsg.includes('undefined')) {
      return {
        success: true,
        available: false,
        needsService: true,
        message: 'Calendar API is not enabled. Please enable it in Apps Script project settings.'
      };
    }

    // Unknown error
    return {
      success: false,
      available: false,
      error: errorMsg,
      message: 'Error checking Calendar API status'
    };
  }
}

/**
 * [CORRECTED VERSION] Fetches today's events from the user's primary calendar.
 * This uses the Advanced Calendar Service to reliably get Google Meet links.
 * @returns {Array} A list of event objects for the frontend.
 */
function getTodaysCalendarEvents() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const now = new Date();
    const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const endOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);

    // Check if Calendar API is available
    if (typeof Calendar === 'undefined') {
      Logger.log('Calendar API not available - Advanced Calendar service not enabled');
      return {
        success: false,
        error: 'CALENDAR_NOT_ENABLED',
        message: 'Calendar integration is not enabled. Enable it in Settings to see your events.'
      };
    }

    // Use the Advanced Calendar service (Calendar.*)
    const eventsList = Calendar.Events.list('primary', {
      timeMin: startOfDay.toISOString(),
      timeMax: endOfDay.toISOString(),
      singleEvents: true,
      orderBy: 'startTime'
    });

    if (!eventsList.items || eventsList.items.length === 0) {
      return { success: true, events: [] };
    }

    // Filter out all-day events (they have a 'date' property instead of 'dateTime')
    const timedEvents = eventsList.items.filter(event => event.start.dateTime);

    const events = timedEvents.map(event => {
      let gmeetLink = null;
      if (event.conferenceData && event.conferenceData.entryPoints) {
        const videoEntryPoint = event.conferenceData.entryPoints.find(ep => ep.entryPointType === 'video');
        if (videoEntryPoint) {
          gmeetLink = videoEntryPoint.uri;
        }
      }

      return {
        id: event.id,
        title: event.summary,
        startTime: event.start.dateTime, // Already in ISO format
        endTime: event.end.dateTime,     // Already in ISO format
        gmeetLink: gmeetLink
      };
    });

    return { success: true, events: events };
  } catch (e) {
    Logger.log(`Error fetching calendar events for user ${Session.getActiveUser().getEmail()}: ${e.toString()}\nStack: ${e.stack}`);

    const errorMsg = e.toString();

    // Handle authorization errors
    if (errorMsg.includes('Authorization') || errorMsg.includes('permission')) {
      return {
        success: false,
        error: 'CALENDAR_AUTH_REQUIRED',
        message: 'Calendar access requires authorization. Please enable calendar integration in Settings.'
      };
    }

    // Handle service not enabled
    if (errorMsg.includes('Calendar') || errorMsg.includes('not found')) {
      return {
        success: false,
        error: 'CALENDAR_NOT_ENABLED',
        message: 'Calendar API is not enabled. Contact your administrator.'
      };
    }

    return {
      success: false,
      error: 'CALENDAR_ERROR',
      message: 'Unable to fetch calendar events'
    };
  }
}

function sanitizeKey_(inputString) { return inputString ? inputString.trim().toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '') : null; }

function getDynamicGreeting(userName) {
  try {
    const now = new Date();
    const hour = now.getHours();
    const dayOfWeek = now.getDay(); // 0 = Sunday, 6 = Saturday
    const month = now.getMonth(); // 0 = January
    const date = now.getDate();
    const firstName = userName ? userName.split(' ')[0] : 'there';

    // Philippine Holidays (with flexible date checking)
    const philippineHolidays = [
      { month: 0, date: 1, greeting: '🎊 Happy New Year' },
      { month: 1, date: 14, greeting: '💝 Happy Valentine\'s Day' },
      { month: 1, date: 25, greeting: '💛 Happy EDSA People Power Anniversary' },
      { month: 3, date: 9, greeting: '🕊️ Happy Araw ng Kagitingan' },
      { month: 4, date: 1, greeting: '👷 Happy Labor Day' },
      { month: 5, date: 12, greeting: '🇵🇭 Happy Independence Day' },
      { month: 7, date: 21, greeting: '🦸 Happy Ninoy Aquino Day' },
      { month: 7, date: 26, greeting: '🦸 Happy National Heroes Day' },
      { month: 10, date: 1, greeting: '🎃 Happy All Saints\' Day' },
      { month: 10, date: 30, greeting: '🦸 Happy Bonifacio Day' },
      { month: 11, date: 8, greeting: '🌟 Happy Feast of the Immaculate Conception' },
      { month: 11, date: 24, greeting: '🎄 Merry Christmas Eve' },
      { month: 11, date: 25, greeting: '🎅 Merry Christmas' },
      { month: 11, date: 26, greeting: '🎁 Happy Boxing Day' },
      { month: 11, date: 30, greeting: '🎊 Happy Rizal Day' },
      { month: 11, date: 31, greeting: '🎆 Happy New Year\'s Eve' }
    ];

    // Check for holidays (including day before for early greetings)
    for (const holiday of philippineHolidays) {
      if (month === holiday.month && date === holiday.date) {
        return { greeting: holiday.greeting, name: firstName };
      }

      // Day before major holidays - properly handles month boundaries
      const tomorrow = new Date(now);
      tomorrow.setDate(tomorrow.getDate() + 1);
      const tomorrowMonth = tomorrow.getMonth();
      const tomorrowDate = tomorrow.getDate();

      if (tomorrowMonth === holiday.month && tomorrowDate === holiday.date) {
        if ([11, 0].includes(holiday.month) && [24, 25, 31, 1].includes(holiday.date)) {
          return { greeting: `${holiday.greeting} tomorrow! 🎉`, name: firstName };
        }
      }
    }

    // Check for special weeks
    if (month === 11 && date >= 16 && date <= 24) {
      return { greeting: '🎄 Merry Christmas season', name: firstName };
    }

    // Day-specific greetings
    const dayGreetings = {
      0: '🌟 Happy Weekend',       // Sunday
      1: '💪 Happy Monday',        // Monday (not shown)
      2: '🚀 Happy Tuesday',       // Tuesday (not shown)
      3: '⚡ Happy Wednesday',     // Wednesday (not shown)
      4: '🎯 Happy Thursday',      // Thursday (not shown)
      5: '🎉 Happy Friday',        // Friday
      6: '🌟 Happy Weekend'        // Saturday
    };

    // Time-based greetings
    let timeGreeting;
    if (hour >= 5 && hour < 12) {
      timeGreeting = 'Good morning';
    } else if (hour >= 12 && hour < 13) {
      timeGreeting = 'Good noon';
    } else if (hour >= 13 && hour < 18) {
      timeGreeting = 'Good afternoon';
    } else if (hour >= 18 && hour < 22) {
      timeGreeting = 'Good evening';
    } else {
      timeGreeting = 'Good night';
    }

    // Combine: "Good morning! Happy Friday" or just "Good morning"
    // Use day greeting on Friday (most exciting) and weekends
    if ([5, 6, 0].includes(dayOfWeek)) {
      return { greeting: `${timeGreeting}! ${dayGreetings[dayOfWeek]}`, name: firstName };
    } else {
      return { greeting: timeGreeting, name: firstName };
    }

  } catch (e) {
    Logger.log(`Error in getDynamicGreeting: ${e.toString()}`);
    return { greeting: 'Hello', name: userName || 'there' };
  }
}
function getUserPreferences_(name) {
  const defaultPrefs = {};
  const userKeyPart = sanitizeKey_(name);
  if (!userKeyPart) return defaultPrefs;

  const scriptProps = PropertiesService.getScriptProperties();
  const propKey = `user_prefs_${userKeyPart}`;
  const jsonString = scriptProps.getProperty(propKey);

  if (jsonString) {
    try {
      return JSON.parse(jsonString);
    } catch (e) {
      Logger.log(`Error parsing preferences for ${name}: ${e.toString()}`);
      return defaultPrefs;
    }
  }
  
  return defaultPrefs;
}
function saveUserPreference(userName, key, value) {
  if (!userName || !key) return { success: false };
  const userKeyPart = sanitizeKey_(userName);
  if (!userKeyPart) return { success: false };
  const scriptProps = PropertiesService.getScriptProperties();
  const propKey = `user_prefs_${userKeyPart}`;
  try {
    const prefs = getUserPreferences_(userName);
    prefs[key] = value;
    scriptProps.setProperty(propKey, JSON.stringify(prefs));

    const automationKeys = ['enableAutoClockOut', 'enableAutoEndBreak', 'enableAutoEndLunch', 'timedActionTriggers'];
    if (automationKeys.includes(key)) {
      ensureAutomationTrigger_();
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
function getSpreadsheetUrl() { try { return SpreadsheetApp.getActiveSpreadsheet().getUrl(); } catch (e) { return "#"; } }
function getProfilePicDataUrlForFileId(fileId) {
  if (!fileId) return { success: false };

  try {
    const file = DriveApp.getFileById(fileId);
    const fileName = file.getName();
    const userEmail = Session.getActiveUser().getEmail();
    const sanitizedEmail = userEmail.replace(/[@.]/g, '_');

    // Security check: Allow users to view:
    // 1. Their own uploaded backgrounds (bg_custom_${email}_)
    // 2. Predefined system backgrounds (don't start with bg_custom_)
    // 3. Admins can view everything
    // Block: Other users' backgrounds (bg_custom_${otherEmail}_)

    const isUserOwnBackground = fileName.startsWith(`bg_custom_${sanitizedEmail}_`);
    const isPredefinedBackground = !fileName.startsWith('bg_custom_');
    const isAdmin = isAdmin_();

    if (!isUserOwnBackground && !isPredefinedBackground && !isAdmin) {
      return { success: false, error: 'Permission denied. You can only view your own backgrounds.' };
    }

    const blob = file.getBlob();
    return {
      success: true,
      dataUrl: `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`
    };
  } catch (e) {
    return { success: false };
  }
}
function getOrCreateTargetSheet_() { const ss = SpreadsheetApp.getActiveSpreadsheet(); const date = new Date(); date.setHours(0, 0, 0, 0); const dayNum = date.getDay() || 7; date.setDate(date.getDate() + 4 - dayNum); const yearStart = new Date(date.getFullYear(), 0, 1); const weekNo = Math.ceil((((date - yearStart) / 86400000) + 1) / 7); const targetSheetName = `Week${weekNo}`; let sheet = ss.getSheetByName(targetSheetName); if (!sheet) { const templateSheet = ss.getSheetByName(TEMPLATE_SHEET_NAME); if (!templateSheet) return null; sheet = ss.insertSheet(targetSheetName, ss.getNumSheets(), { template: templateSheet }); } return sheet; }
function findNameDataForDate_(sheet, targetDate) { if (!sheet || !(targetDate instanceof Date)) return null; const targetMillis = Date.UTC(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate()); try { const dateRanges = sheet.getRangeList(DATE_NAME_MAPPING.map(m => `A${m.dateRow}`)).getRanges(); const dateValues = dateRanges.map(r => r.getValue()); const mapIndex = dateValues.findIndex(val => val instanceof Date && Date.UTC(val.getFullYear(), val.getMonth(), val.getDate()) === targetMillis); if (mapIndex === -1) return []; const { nameStartRow, nameEndRow } = DATE_NAME_MAPPING[mapIndex]; const nameVals = sheet.getRange(nameStartRow, 1, nameEndRow - nameStartRow + 1, 1).getValues(); return nameVals.map((r, i) => ({ name: r[0].trim(), row: nameStartRow + i })).filter(item => item.name); } catch (e) { return null; } }
function getTimeEntriesForUser_(sheet, row) { if (!sheet || !row) return null; try { const v = sheet.getRange(row, 2, 1, 6).getDisplayValues()[0]; return { clockIn: v[0]||null, breakOut: v[1]||null, breakIn: v[2]||null, lunchOut: v[3]||null, lunchIn: v[4]||null, clockOut: v[5]||null }; } catch(e) { return null; } }
function determineStatusFromTimes(times) { const has = (key) => times && times[key] && String(times[key]).trim() !== "" && String(times[key]).trim() !== "--"; if (has('clockOut')) return 'Offline'; if (has('lunchIn')) return 'Working'; if (has('lunchOut')) return 'OnLunch'; if (has('breakIn')) return 'Working'; if (has('breakOut')) return 'OnBreak'; if (has('clockIn')) return 'Working'; return 'Offline'; }
function calculateCurrentTotalHoursSeconds_(sheet, row) { if (!sheet || !row) return 0; try { const range = sheet.getRange(row, 2, 1, 6); const rawVals = range.getValues()[0]; const getMs = v => (v instanceof Date) ? v.getTime() : null; const clockInMs = getMs(rawVals[0]); if (!clockInMs) return 0; const status = determineStatusFromTimes(getTimeEntriesForUser_(sheet, row)); let effectiveEndMs; if (getMs(rawVals[5])) { effectiveEndMs = getMs(rawVals[5]); } else if (status === 'Working') { effectiveEndMs = new Date().getTime(); } else if (status === 'OnBreak') { effectiveEndMs = getMs(rawVals[1]); } else if (status === 'OnLunch') { effectiveEndMs = getMs(rawVals[3]); } else { effectiveEndMs = [getMs(rawVals[4]), getMs(rawVals[2]), clockInMs].find(t => t) || clockInMs; } let durationMs = effectiveEndMs - clockInMs; if (getMs(rawVals[1]) && getMs(rawVals[2])) { durationMs -= (getMs(rawVals[2]) - getMs(rawVals[1])); } if (getMs(rawVals[3]) && getMs(rawVals[4])) { durationMs -= (getMs(rawVals[4]) - getMs(rawVals[3])); } return Math.round(Math.max(0, durationMs) / 1000); } catch (e) { return 0; } }

function updateTimeEntry(row, actionKey, newTimeString) {
  if (!isAdmin_()) {
    return { success: false, error: 'Permission denied.' };
  }
  try {
    const sheet = getOrCreateTargetSheet_();
    if (!sheet) throw new Error("Could not access the timesheet.");
    
    const column = COLUMN_MAP[actionKey];
    if (!column) throw new Error("Invalid action key provided.");
    
    if (!newTimeString || newTimeString.trim() === '') {
      sheet.getRange(row, column.index).clearContent();
      SpreadsheetApp.flush();
      recalculateHours(row, sheet);
      return { success: true, message: `Cleared ${column.name}` };
    }

    let targetRow = row;
    let userName = sheet.getRange(row, 1).getValue().trim();
    let newTime;

    if (newTimeString.includes('/')) {
      const dateStr = newTimeString.split(' ')[0];
      const targetDate = new Date(dateStr);
      if (isNaN(targetDate.getTime())) { throw new Error("Invalid date provided in the string."); }
      
      const userEntryOnDate = findNameDataForDate_(sheet, targetDate)?.find(u => u.name === userName);
      if (userEntryOnDate && userEntryOnDate.row) {
        targetRow = userEntryOnDate.row;
      } else {
        throw new Error(`Could not find "${userName}" on the schedule for ${targetDate.toLocaleDateString()}.`);
      }
      
      try {
        newTime = Utilities.parseDate(newTimeString, SCRIPT_TIMEZONE, "M/d/yyyy HH:mm:ss");
      } catch (e) {
        throw new Error("Invalid format. Please use 'M/d/yyyy HH:mm:ss' (24h).");
      }

    } else {
      const rowDateValue = sheet.getRange(findDateRowForRow_(targetRow), 1).getValue();
      const targetDate = new Date(rowDateValue);
      const datePart = (targetDate.getMonth() + 1) + '/' + targetDate.getDate() + '/' + targetDate.getFullYear();
      const combinedString = datePart + ' ' + newTimeString;
      
      try {
        newTime = new Date(combinedString);
        if (isNaN(newTime.getTime())) {
           throw new Error("Could not parse time.");
        }
      } catch (e) {
        throw new Error("Invalid time format. Please use a valid time like '4:15:30 PM' or '16:15:30'.");
      }
    }

    // VALIDATION: Check logical consistency of time entries
    const allValues = sheet.getRange(targetRow, 2, 1, 6).getValues()[0];
    const timeOrder = ['clockIn', 'breakOut', 'breakIn', 'lunchOut', 'lunchIn', 'clockOut'];
    const columnIndexes = {
      'clockIn': 0, 'breakOut': 1, 'breakIn': 2,
      'lunchOut': 3, 'lunchIn': 4, 'clockOut': 5
    };

    // Create a temporary array with the new value inserted
    const tempValues = [...allValues];
    tempValues[columnIndexes[actionKey]] = newTime;

    // Validate chronological order of non-empty entries
    let lastTime = null;
    for (let i = 0; i < timeOrder.length; i++) {
      const currentValue = tempValues[columnIndexes[timeOrder[i]]];
      if (currentValue instanceof Date) {
        if (lastTime && currentValue.getTime() < lastTime.getTime()) {
          throw new Error(
            `Invalid time order: ${COLUMN_MAP[timeOrder[i]].name} (${currentValue.toLocaleString()}) cannot be before ${lastTimeName} (${lastTime.toLocaleString()})`
          );
        }
        lastTime = currentValue;
        var lastTimeName = COLUMN_MAP[timeOrder[i]].name;
      }
    }

    // Validate: Break In must come after Break Out
    if (tempValues[columnIndexes['breakOut']] instanceof Date &&
        tempValues[columnIndexes['breakIn']] instanceof Date &&
        tempValues[columnIndexes['breakIn']].getTime() <= tempValues[columnIndexes['breakOut']].getTime()) {
      throw new Error('Break In must be after Break Out');
    }

    // Validate: Lunch In must come after Lunch Out
    if (tempValues[columnIndexes['lunchOut']] instanceof Date &&
        tempValues[columnIndexes['lunchIn']] instanceof Date &&
        tempValues[columnIndexes['lunchIn']].getTime() <= tempValues[columnIndexes['lunchOut']].getTime()) {
      throw new Error('Lunch In must be after Lunch Out');
    }

    // All validations passed - save the value
    sheet.getRange(targetRow, column.index).setValue(newTime).setNumberFormat(TIME_CELL_FORMAT);
    SpreadsheetApp.flush();
    recalculateHours(targetRow, sheet);

    return { success: true, message: `Updated ${column.name}` };
  } catch (e) {
    Logger.log(`Admin update error: ${e.stack}`);
    return { success: false, error: e.message };
  }
}

function findDateRowForRow_(row) {
  for (let i = DATE_NAME_MAPPING.length - 1; i >= 0; i--) {
    const mapping = DATE_NAME_MAPPING[i];
    if (row >= mapping.nameStartRow) {
      return mapping.dateRow;
    }
  }
  return 1;
}

function deleteTimeEntry(row, actionKey) { if (!isAdmin_()) return { success: false, error: 'Permission denied.' }; try { const sheet = getOrCreateTargetSheet_(); if (!sheet) throw new Error("Could not access timesheet."); const column = COLUMN_MAP[actionKey]; if (!column) throw new Error("Invalid action key."); sheet.getRange(row, column.index).clearContent(); SpreadsheetApp.flush(); recalculateHours(row, sheet); return { success: true }; } catch (e) { return { success: false, error: e.message }; } }
function recalculateHours(row, sheet) { try { const targetSheet = sheet || getOrCreateTargetSheet_(); if (!targetSheet) throw new Error("Could not access timesheet."); const totalHours = calculateCurrentTotalHoursSeconds_(targetSheet, row) / 3600; const decimalHoursFormat = "0.00"; targetSheet.getRange(row, COLUMN_MAP.hours.index).setValue(totalHours > 0.001 ? totalHours : "").setNumberFormat(decimalHoursFormat); SpreadsheetApp.flush(); return { success: true, newTotalHours: totalHours.toFixed(2) }; } catch(e) { return { success: false, error: e.message }; } }
function saveProfilePicture(userName, base64Data, mimeType) { if (!isAdmin_()) return { success: false, error: 'Permission denied.' }; try { let folders = DriveApp.getFoldersByName(ASSET_FOLDER_NAME); let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(ASSET_FOLDER_NAME); const sanitizedKey = sanitizeKey_(userName); if (!sanitizedKey) return { success: false, error: "Invalid user name." }; const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, `profile_${sanitizedKey}.png`); const propKey = `profilePic_name_${sanitizedKey}`; const oldFileId = PropertiesService.getScriptProperties().getProperty(propKey); if (oldFileId) { try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (e) {} } const newFile = folder.createFile(blob); PropertiesService.getScriptProperties().setProperty(propKey, newFile.getId()); return { success: true, fileId: newFile.getId() }; } catch (e) { return { success: false, error: e.message }; } }
function findUserByName_(userName) { const allUsersResult = getAllUsersStatus(); if (allUsersResult.success) { return allUsersResult.users.find(user => user.name === userName); } return null; }
function generateWeekSheet(weekNumber) {
  if (!isAdmin_()) { return { success: false, error: 'Permission denied.' }; }
  try {
    const year = new Date().getFullYear();
    const sheetName = `Week${weekNumber}`;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss.getSheetByName(sheetName)) { throw new Error(`Sheet "${sheetName}" already exists.`); }
    const templateSheet = ss.getSheetByName(TEMPLATE_SHEET_NAME);
    if (!templateSheet) { throw new Error(`Template sheet named "${TEMPLATE_SHEET_NAME}" not found.`); }
    const newSheet = templateSheet.copyTo(ss).setName(sheetName);
    ss.setActiveSheet(newSheet);
    const janFirst = new Date(year, 0, 1);
    const daysOffset = (weekNumber - 1) * 7 - janFirst.getDay() + 1;
    const firstDayOfWeek = new Date(janFirst.getFullYear(), janFirst.getMonth(), janFirst.getDate() + daysOffset);
    const headerFont = "Inter"; const headerFontSize = 11; const headerBackgroundColor = "#000000"; const headerFontColor = "#ffffff"; const headerFontWeight = "bold";
    DATE_NAME_MAPPING.forEach((mapping, index) => {
        const date = new Date(firstDayOfWeek);
        date.setDate(date.getDate() + index);
        newSheet.getRange(`A${mapping.dateRow}`).setValue(date).setNumberFormat('mmmm" "d", "yyyy').setBackground(headerBackgroundColor).setFontColor(headerFontColor).setFontFamily(headerFont).setFontSize(headerFontSize).setFontWeight(headerFontWeight);
    });
    return { success: true, message: `Successfully created sheet: ${sheetName}` };
  } catch(e) {
    Logger.log(`Error generating week sheet: ${e.stack}`);
    return { success: false, error: e.message };
  }
}
function recalculateHoursForDate(userName, dateString) {
  if (!isAdmin_()) { return { success: false, error: 'Permission denied.' }; }
  try {
    const targetDate = new Date(dateString);
    const sheet = getOrCreateTargetSheet_(); 
    if (!sheet) throw new Error("Could not access the timesheet.");
    const userEntry = findNameDataForDate_(sheet, targetDate)?.find(u => u.name === userName);
    if (!userEntry || !userEntry.row) { throw new Error(`Could not find entry for ${userName} on ${targetDate.toLocaleDateString()}.`); }
    const result = recalculateHours(userEntry.row, sheet);
    if (!result.success) throw new Error(result.error);
    return { success: true, newTotalHours: result.newTotalHours };
  } catch (e) {
    Logger.log(`Error recalculating hours for date: ${e.stack}`);
    return { success: false, error: e.message };
  }
}
function getBackgroundImages() {
  try {
    const folders = DriveApp.getFoldersByName(ASSET_FOLDER_NAME);
    if (!folders.hasNext()) {
      return { success: true, images: [] };
    }
    const folder = folders.next();
    const images = [];
    const userEmail = Session.getActiveUser().getEmail();
    const sanitizedEmail = userEmail.replace(/[@.]/g, '_'); // Sanitize email for filename matching

    // Search for files that belong to this user or are predefined backgrounds
    const allFiles = folder.getFiles();
    while (allFiles.hasNext()) {
      const file = allFiles.next();
      const fileName = file.getName();

      // Include predefined backgrounds (bg1-bg5) for everyone
      if (/^bg[1-5]$/.test(fileName)) {
        images.push({ name: fileName, id: file.getId() });
      }
      // Include custom backgrounds only if they belong to this user
      else if (fileName.startsWith(`bg_custom_${sanitizedEmail}_`)) {
        images.push({ name: fileName, id: file.getId() });
      }
    }

    return { success: true, images: images };
  } catch(e) {
    Logger.log(`Error in getBackgroundImages: ${e.toString()}`);
    return { success: false, error: e.message };
  }
}

function saveBackgroundImage(base64Data, mimeType) {
  try {
    let folders = DriveApp.getFoldersByName(ASSET_FOLDER_NAME);
    let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(ASSET_FOLDER_NAME);

    const userEmail = Session.getActiveUser().getEmail();
    const sanitizedEmail = userEmail.replace(/[@.]/g, '_'); // Sanitize email for filename
    const timestamp = new Date().getTime();
    const fileName = `bg_custom_${sanitizedEmail}_${timestamp}.png`;
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);

    const newFile = folder.createFile(blob);
    return { success: true, fileId: newFile.getId() };
  } catch (e) {
    Logger.log(`Error in saveBackgroundImage: ${e.toString()}`);
    return { success: false, error: e.message };
  }
}

function deleteBackgroundImage(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const fileName = file.getName();
    const userEmail = Session.getActiveUser().getEmail();
    const sanitizedEmail = userEmail.replace(/[@.]/g, '_');

    // Security check: Only allow users to delete their own backgrounds
    // (or allow admins to delete any)
    if (!fileName.startsWith(`bg_custom_${sanitizedEmail}_`) && !isAdmin_()) {
      return { success: false, error: 'Permission denied. You can only delete your own backgrounds.' };
    }

    file.setTrashed(true);
    return { success: true };
  } catch (e) {
    Logger.log(`Error in deleteBackgroundImage: ${e.toString()}`);
    return { success: false, error: e.message };
  }
}

function saveBackgroundPreference(selection) {
    if (!isAdmin_()) { return { success: false, error: 'Permission denied.' }; }
    try {
        const propKey = 'backgroundSelection';
        if (selection === 'default') {
            PropertiesService.getScriptProperties().deleteProperty(propKey);
        } else {
            PropertiesService.getScriptProperties().setProperty(propKey, selection);
        }
        return { success: true };
    } catch(e) {
        return { success: false, error: e.message };
    }
}

function logAutomationEvent_(userName, action, reason) {
  const scriptProps = PropertiesService.getScriptProperties();
  const newEntry = { 
    timestamp: new Date().toISOString(), 
    userName: userName, 
    action: action, 
    reason: reason 
  };

  const logString = scriptProps.getProperty(LOG_KEYS.AUTOMATION_LOG);
  let log = logString ? JSON.parse(logString) : [];
  log.unshift(newEntry);
  if (log.length > MAX_LOG_ENTRIES) {
    log = log.slice(0, MAX_LOG_ENTRIES);
  }
  scriptProps.setProperty(LOG_KEYS.AUTOMATION_LOG, JSON.stringify(log));

  const recentEventsString = scriptProps.getProperty(LOG_KEYS.RECENT_AUTOMATION_EVENTS);
  let recentEvents = recentEventsString ? JSON.parse(recentEventsString) : [];
  recentEvents.push(newEntry);
  scriptProps.setProperty(LOG_KEYS.RECENT_AUTOMATION_EVENTS, JSON.stringify(recentEvents));
}

function getAutomationLogs() {
  if (!isSuperAdmin_()) { return { success: false, error: 'Permission denied.' }; }
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const logString = scriptProps.getProperty(LOG_KEYS.AUTOMATION_LOG);
    const allLogs = logString ? JSON.parse(logString) : [];

    const twentyFourHoursAgo = new Date(Date.now() - 24 * 60 * 60 * 1000);
    
    const recentLogs = allLogs.filter(log => new Date(log.timestamp) >= twentyFourHoursAgo);

    if (recentLogs.length < allLogs.length) {
      scriptProps.setProperty(LOG_KEYS.AUTOMATION_LOG, JSON.stringify(recentLogs));
      Logger.log(`AUTOMATION: Pruned ${allLogs.length - recentLogs.length} old log entries.`);
    }

    return { success: true, logs: recentLogs };
  } catch (e) {
    Logger.log(`Error in getAutomationLogs: ${e.stack}`);
    return { success: false, error: e.message };
  }
}

function clearAutomationLogs() {
  if (!isSuperAdmin_()) { return { success: false, error: 'Permission denied.' }; }
  try {
    PropertiesService.getScriptProperties().deleteProperty(LOG_KEYS.AUTOMATION_LOG);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getRecentAutomationEvents() {
    if (!isSuperAdmin_()) { return { success: false, events: [] }; }
    try {
        const scriptProps = PropertiesService.getScriptProperties();
        const eventsString = scriptProps.getProperty(LOG_KEYS.RECENT_AUTOMATION_EVENTS);
        if (eventsString) {
            scriptProps.deleteProperty(LOG_KEYS.RECENT_AUTOMATION_EVENTS);
            return { success: true, events: JSON.parse(eventsString) };
        }
        return { success: true, events: [] };
    } catch (e) {
        return { success: false, events: [] };
    }
}


function ensureAutomationTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  const existingTrigger = triggers.find(t => t.getHandlerFunction() === 'runAutomations');

  if (!existingTrigger) {
    ScriptApp.newTrigger('runAutomations')
      .timeBased()
      .everyMinutes(1)
      .create();
    Logger.log('Automation trigger created to run every minute.');
    return true;
  }
  return false;
}

function createAutomationTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'runAutomations') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  ScriptApp.newTrigger('runAutomations')
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log('Automation trigger created to run every minute.');
}

function checkAutomationStatus() {
  if (!isSuperAdmin_()) {
    return { success: false, error: 'Permission denied.' };
  }

  try {
    const triggers = ScriptApp.getProjectTriggers();
    const automationTrigger = triggers.find(t => t.getHandlerFunction() === 'runAutomations');

    if (automationTrigger) {
      return {
        success: true,
        status: 'active',
        message: 'Automation trigger is active and running every minute.',
        triggerInfo: {
          handlerFunction: automationTrigger.getHandlerFunction(),
          triggerSource: automationTrigger.getTriggerSource().toString(),
          eventType: automationTrigger.getEventType().toString()
        }
      };
    } else {
      return {
        success: true,
        status: 'inactive',
        message: 'Automation trigger is NOT active. Creating it now...'
      };
    }
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function debugTimedActions(userName) {
  if (!isSuperAdmin_()) {
    return { success: false, error: 'Permission denied.' };
  }

  try {
    const now = new Date();
    const prefs = getUserPreferences_(userName);
    const profileResult = getUserProfileData(userName);

    if (!profileResult.success) {
      return { success: false, error: 'Could not get user profile data.' };
    }

    const profileData = profileResult.data;
    const todayStr = now.toISOString().slice(0, 10);
    const scriptProps = PropertiesService.getScriptProperties();

    let debugInfo = {
      success: true,
      currentTime: now.toLocaleTimeString(),
      userName: userName,
      userStatus: profileData.status,
      timedActionTriggers: null,
      triggerStates: []
    };

    if (!prefs.timedActionTriggers) {
      debugInfo.message = 'No timed action triggers found for this user.';
      return debugInfo;
    }

    const triggers = JSON.parse(prefs.timedActionTriggers);
    debugInfo.timedActionTriggers = triggers;

    triggers.forEach(trigger => {
      const [triggerHour, triggerMinute] = trigger.time.split(':').map(Number);
      const triggerTime = new Date(now);
      triggerTime.setHours(triggerHour, triggerMinute, 0, 0);
      const windowEnd = new Date(triggerTime.getTime() + 60 * 60 * 1000);

      const keyPart = sanitizeKey_(userName) + '_' + trigger.action + '_' + trigger.time.replace(':', '');
      const hasRunKey = `timedActionRan_${keyPart}_${todayStr}`;
      const randomTimeKey = `timedActionRandomTime_${keyPart}_${todayStr}`;

      const hasRun = scriptProps.getProperty(hasRunKey);
      const randomTime = scriptProps.getProperty(randomTimeKey);

      const triggerState = {
        action: trigger.action,
        time: trigger.time,
        enabled: trigger.enabled,
        triggerTime: triggerTime.toLocaleTimeString(),
        windowEnd: windowEnd.toLocaleTimeString(),
        inWindow: (now >= triggerTime && now < windowEnd),
        hasRunToday: !!hasRun,
        randomExecutionTime: randomTime ? new Date(parseInt(randomTime)).toLocaleTimeString() : 'Not generated yet',
        shouldExecuteNow: false
      };

      if (randomTime) {
        const targetTime = parseInt(randomTime);
        triggerState.shouldExecuteNow = (now.getTime() >= targetTime && !hasRun && trigger.enabled);
      }

      debugInfo.triggerStates.push(triggerState);
    });

    return debugInfo;
  } catch (e) {
    return { success: false, error: e.message, stack: e.stack };
  }
}

function runAutomations() {
  const now = new Date();
  Logger.log(`========== AUTOMATION RUN START: ${now.toLocaleString()} ==========`);

  // Check if today is a weekend (0 = Sunday, 6 = Saturday)
  const dayOfWeek = now.getDay();
  const isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);

  if (isWeekend) {
    Logger.log(`AUTOMATION: Today is a weekend (day ${dayOfWeek}). Checking user preferences for weekend automation...`);
  }

  const allUsersResult = getAllUsersStatus();
  if (!allUsersResult.success) {
    Logger.log("AUTOMATION: Failed to get user statuses. Aborting.");
    return;
  }

  Logger.log(`AUTOMATION: Processing ${allUsersResult.users.length} user(s)`);
  const todayStr = now.toISOString().slice(0, 10);

  allUsersResult.users.forEach(user => {
    try {
      const prefs = getUserPreferences_(user.name);
      if (!prefs) return;

      // Skip automation on weekends if user has enabled skipWeekendAutomation
      if (isWeekend && prefs[PREF_KEYS.SKIP_WEEKEND_AUTOMATION] === 'true') {
        Logger.log(`AUTOMATION: Skipping ${user.name} - weekend automation disabled by user preference`);
        return;
      }

      const profileResult = getUserProfileData(user.name);
      if (!profileResult.success) return;
      const profileData = profileResult.data;

      handleAutoClockOut_(profileData, prefs, todayStr);
      handleAutoEndTimers_(profileData, prefs);
      handleTimedActions_(profileData, prefs, now);

    } catch (e) {
      Logger.log(`AUTOMATION: Error processing user ${user.name}: ${e.stack}`);
    }
  });
}

function handleAutoClockOut_(profileData, prefs, todayStr) {
  if (prefs.enableAutoClockOut !== 'true' || profileData.status !== 'Working') return;

  const durationHours = parseFloat(prefs.autoClockOutDuration) || 8;
  const workedHours = profileData.totalSecondsToday / 3600;

  if (workedHours >= durationHours && workedHours < durationHours + 1) {
    const hasRunKey = `autoClockOutRan_${sanitizeKey_(profileData.name)}_${todayStr}`;
    if (PropertiesService.getScriptProperties().getProperty(hasRunKey)) return;

    const minutesIntoWindow = Math.floor((workedHours - durationHours) * 60);
    const minutesRemaining = Math.max(1, 60 - minutesIntoWindow);

    if (Math.floor(Math.random() * minutesRemaining) === 0) {
      const actionTimestamp = new Date();
      actionTimestamp.setSeconds(Math.floor(Math.random() * 60));

      Logger.log(`AUTOMATION: Firing random auto-clock-out for ${profileData.name} at ${workedHours.toFixed(2)} hours.`);
      processTimeEntry(profileData.name, 'clockOut', actionTimestamp);
      logAutomationEvent_(profileData.name, 'Auto Clock Out', `Reached ${durationHours.toFixed(1)} working hours.`);
      PropertiesService.getScriptProperties().setProperty(hasRunKey, 'true');
    }
  }
}

function handleAutoEndTimers_(profileData, prefs) {
  const { status, timeEntries, name, actualDataRowForStatus } = profileData;
  if (!actualDataRowForStatus) return;

  const sheet = getOrCreateTargetSheet_();
  const dateRow = findDateRowForRow_(actualDataRowForStatus);
  const dateOfRow = sheet.getRange(dateRow, 1).getValue();

  const createFullDate = (timeString) => {
    if (!timeString || timeString === '--' || !dateOfRow) return null;
    try {
      const datePart = (dateOfRow.getMonth() + 1) + '/' + dateOfRow.getDate() + '/' + dateOfRow.getFullYear();
      return new Date(datePart + ' ' + timeString);
    } catch (e) {
      Logger.log(`AUTOMATION: Could not parse time string "${timeString}" for user ${name}`);
      return null;
    }
  };

  if (prefs.enableAutoEndBreak === 'true' && status === 'OnBreak') {
    const durationMinutes = parseInt(prefs.autoEndBreakDuration, 10) || 15;
    const breakOutTime = createFullDate(timeEntries.breakOut);
    if (breakOutTime) {
       const minutesOnBreak = (new Date().getTime() - breakOutTime.getTime()) / 60000;
       if (minutesOnBreak >= durationMinutes) {
         Logger.log(`AUTOMATION: Auto-ending break for ${name}.`);
         processTimeEntry(name, 'breakIn');
         logAutomationEvent_(name, 'Auto End Break', `Exceeded ${durationMinutes} minute duration.`);
       }
    }
  }

  if (prefs.enableAutoEndLunch === 'true' && status === 'OnLunch') {
    const durationMinutes = parseInt(prefs.autoEndLunchDuration, 10) || 60;
    const lunchOutTime = createFullDate(timeEntries.lunchOut);
    if (lunchOutTime) {
       const minutesOnLunch = (new Date().getTime() - lunchOutTime.getTime()) / 60000;
       if (minutesOnLunch >= durationMinutes) {
         Logger.log(`AUTOMATION: Auto-ending lunch for ${name}.`);
         processTimeEntry(name, 'lunchIn');
         logAutomationEvent_(name, 'Auto End Lunch', `Exceeded ${durationMinutes} minute duration.`);
       }
    }
  }
}

function handleTimedActions_(profileData, prefs, now) {
  if (!prefs.timedActionTriggers) {
    Logger.log(`AUTOMATION: No timed action triggers for ${profileData.name}`);
    return;
  }
  try {
    const triggers = JSON.parse(prefs.timedActionTriggers);
    if (!Array.isArray(triggers)) {
      Logger.log(`AUTOMATION: Timed action triggers not an array for ${profileData.name}`);
      return;
    }

    Logger.log(`AUTOMATION: Checking ${triggers.length} timed action(s) for ${profileData.name} at ${now.toLocaleTimeString()}`);

    const todayStr = now.toISOString().slice(0, 10);
    const scriptProps = PropertiesService.getScriptProperties();

    triggers.forEach(trigger => {
      if (!trigger.enabled || !trigger.time) {
        Logger.log(`AUTOMATION: Skipping disabled or malformed trigger for ${profileData.name}: ${JSON.stringify(trigger)}`);
        return;
      }

      const [triggerHour, triggerMinute] = trigger.time.split(':').map(Number);

      // Create the trigger time for today
      const triggerTime = new Date(now);
      triggerTime.setHours(triggerHour, triggerMinute, 0, 0);

      // Create the window end time (trigger time + 60 minutes)
      const windowEnd = new Date(triggerTime.getTime() + 60 * 60 * 1000);

      Logger.log(`AUTOMATION: ${profileData.name} ${trigger.action} - Trigger: ${triggerTime.toLocaleTimeString()}, Window: ${triggerTime.toLocaleTimeString()} to ${windowEnd.toLocaleTimeString()}, Current: ${now.toLocaleTimeString()}`);

      // Check if we're within the execution window
      if (now < triggerTime || now >= windowEnd) {
        Logger.log(`AUTOMATION: Outside execution window for ${profileData.name} ${trigger.action}`);
        return;
      }

      const keyPart = sanitizeKey_(profileData.name) + '_' + trigger.action + '_' + trigger.time.replace(':', '');
      const hasRunKey = `timedActionRan_${keyPart}_${todayStr}`;
      if (scriptProps.getProperty(hasRunKey)) {
        Logger.log(`AUTOMATION: Already ran today for ${profileData.name} ${trigger.action}`);
        return;
      }

      // Check if the action was already performed manually today
      const actionColumn = COLUMN_MAP[trigger.action];
      if (!actionColumn) {
        Logger.log(`AUTOMATION: Invalid action type ${trigger.action} for ${profileData.name}`);
        return;
      }

      const timeEntries = profileData.timeEntries;
      const actionAlreadyPerformed = timeEntries && timeEntries[trigger.action] &&
                                     timeEntries[trigger.action] !== '--' &&
                                     timeEntries[trigger.action].trim() !== '';

      if (actionAlreadyPerformed) {
        Logger.log(`AUTOMATION: Action ${trigger.action} already performed manually today for ${profileData.name} at ${timeEntries[trigger.action]}. Skipping automation.`);
        scriptProps.setProperty(hasRunKey, 'true'); // Mark as complete to prevent future checks
        return;
      }

      // Get or generate the random execution time for this trigger
      const randomTimeKey = `timedActionRandomTime_${keyPart}_${todayStr}`;
      let randomExecutionTime = scriptProps.getProperty(randomTimeKey);

      if (!randomExecutionTime) {
        // First time entering window - pick a random minute (0-59) and second (0-59)
        const randomMinutes = Math.floor(Math.random() * 60);
        const randomSeconds = Math.floor(Math.random() * 60);
        const randomTime = new Date(triggerTime.getTime() + randomMinutes * 60000 + randomSeconds * 1000);
        randomExecutionTime = randomTime.getTime().toString();
        scriptProps.setProperty(randomTimeKey, randomExecutionTime);
        Logger.log(`AUTOMATION: Generated random execution time for ${profileData.name} ${trigger.action}: ${new Date(parseInt(randomExecutionTime)).toLocaleTimeString()} (${randomMinutes}m ${randomSeconds}s into window)`);
      } else {
        Logger.log(`AUTOMATION: Using existing random time for ${profileData.name} ${trigger.action}: ${new Date(parseInt(randomExecutionTime)).toLocaleTimeString()}`);
      }

      const targetExecutionTime = parseInt(randomExecutionTime);

      // Check if it's time to execute (we've passed the random execution time)
      if (now.getTime() < targetExecutionTime) {
        Logger.log(`AUTOMATION: Not yet time to execute ${profileData.name} ${trigger.action}. Waiting until ${new Date(targetExecutionTime).toLocaleTimeString()}`);
        return;
      }

      Logger.log(`AUTOMATION: Time reached for ${profileData.name} ${trigger.action}. Checking status...`);

      const { status, name } = profileData;
      const action = trigger.action;
      let canExecute = false;
      let requiredStatus = '';

      switch (action) {
          case 'clockIn':
              canExecute = (status === 'Offline');
              requiredStatus = 'Offline';
              break;
          case 'breakOut':
          case 'lunchOut':
          case 'clockOut':
              canExecute = (status === 'Working');
              requiredStatus = 'Working';
              break;
          case 'breakIn':
              canExecute = (status === 'OnBreak');
              requiredStatus = 'OnBreak';
              break;
          case 'lunchIn':
              canExecute = (status === 'OnLunch');
              requiredStatus = 'OnLunch';
              break;
      }

      Logger.log(`AUTOMATION: Status check for ${name} ${action} - Current: ${status}, Required: ${requiredStatus}, Can execute: ${canExecute}`);

      if (canExecute) {
          const actionTimestamp = new Date(targetExecutionTime);
          const minutesIntoWindow = Math.floor((actionTimestamp.getTime() - triggerTime.getTime()) / 60000);

          Logger.log(`AUTOMATION: Firing timed action "${action}" for ${name}. Scheduled: ${trigger.time}, Executing at pre-determined random time: ${actionTimestamp.toLocaleTimeString()} (${minutesIntoWindow} min into window).`);
          processTimeEntry(name, action, actionTimestamp);

          const prettyAction = action.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase());
          logAutomationEvent_(name, `Timed Action: ${prettyAction}`, `Trigger set for ${trigger.time}, executed at ${actionTimestamp.toLocaleTimeString()} (${minutesIntoWindow}m ${actionTimestamp.getSeconds()}s into window).`);

          scriptProps.setProperty(hasRunKey, 'true');
          scriptProps.deleteProperty(randomTimeKey); // Clean up the random time
      } else {
          Logger.log(`AUTOMATION: Waiting to execute timed action "${action}" for ${name}. Required status not met. Scheduled random time: ${new Date(targetExecutionTime).toLocaleTimeString()}, Current status: ${status}.`);
      }
    });
  } catch(e) {
    Logger.log(`AUTOMATION: Could not parse or execute timed actions for ${profileData.name}: ${e.toString()}`);
  }
}
