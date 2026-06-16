/**
 * =========================================================================================
 * SECTION 0: DYNAMIC SETTINGS & ROLES READER
 * =========================================================================================
 */

function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Settings");
  if (!sheet) {
    return {
      sites: ["Site#1", "Site#2", "Site#3", "Site#4"],
      events: ["Event1", "Event2", "Event3"],
      schedules: {},
      siteGroups: {},
      appTitle: "Event Check-In App",
      appVersion: "v1.0.0"
    };
  }

  const data = sheet.getDataRange().getValues();
  const settings = { sites: [], events: [], schedules: {}, siteGroups: {}, appTitle: "Event Check-In App", appVersion: "v1.0.0" };

  // 1. Extract simple key‑value pairs (Sites, Event Types, App Version, App Title)
  for (let i = 0; i < data.length; i++) {
    const key = data[i][0] ? data[i][0].toString().trim() : "";
    const val = data[i][1] ? data[i][1].toString().trim() : "";
    if (key === "Event Sites") {
      settings.sites = val.split(",").map(s => s.trim()).filter(s => s);
    } else if (key === "Event Types") {
      settings.events = val.split(",").map(s => s.trim()).filter(s => s);
    } else if (key === "App Version") {
      settings.appVersion = val;
    } else if (key === "App Title") {
      settings.appTitle = val || "Event Check-In App";
    } else if (key === "Site Groups") {
      // Parse each non‑empty row after the key, format: "MainBooth = Site1, Site2, Site3, Site4"
      for (let j = i + 1; j < data.length; j++) {
        const line = data[j][1] ? data[j][1].toString().trim() : "";
        if (!line) break;
        const parts = line.split("=");
        if (parts.length === 2) {
          const groupName = parts[0].trim().toUpperCase();
          const sitesList = parts[1].split(",").map(s => s.trim().toUpperCase()).filter(s => s);
          if (groupName && sitesList.length > 0) {
            settings.siteGroups[groupName] = sitesList;
          }
        }
      }
    }
  }

  // 2. Find schedule table header ("Location / Group")
  let headerRow = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === "Location / Group") {
      headerRow = i;
      break;
    }
  }
  if (headerRow !== -1) {
    for (let i = headerRow + 1; i < data.length; i++) {
      const location = data[i][0] ? data[i][0].toString().trim() : "";
      if (!location) break;
      const eventType = data[i][1] ? data[i][1].toString().trim() : "";
      const startHour = parseInt(data[i][2], 10);
      const startMinute = parseInt(data[i][3], 10);
      if (isNaN(startHour) || isNaN(startMinute)) continue;
      if (!settings.schedules[location]) settings.schedules[location] = {};
      settings.schedules[location][eventType] = { hour: startHour, minute: startMinute };
    }
  }

  return settings;
}

/**
 * Reads the Roles sheet and returns expanded permissions for an email.
 * Columns: Full Name, Email, Role, Site, Group, Permission Level
 * "Permission Level" contains the list of allowed events (comma‑separated).
 * Returns: { role, allowedSites, allowedgroup, permissionLevel, allowedEvents } or null
 */
function getRolePermissions(email) {
  if (!email) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Roles");
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  const header = data[0];
  const colMap = {};
  header.forEach((h, idx) => {
    let key = h.toString().trim().toLowerCase();
    key = key.replace(/:+$/, "").trim();
    if (key === "email") colMap.email = idx;
    else if (key === "role") colMap.role = idx;
    else if (key === "site") colMap.site = idx;
    else if (key === "group") colMap.group = idx;
    else if (key === "permission level") colMap.permissionLevel = idx;
  });

  if (colMap.email === undefined || colMap.role === undefined || colMap.site === undefined) return null;

  const cleanEmail = email.toLowerCase().trim();
  const settings = getSettings();

  for (let i = 1; i < data.length; i++) {
    const rowEmail = data[i][colMap.email] ? data[i][colMap.email].toString().trim().toLowerCase() : "";
    if (rowEmail === cleanEmail) {
      const rawRole = data[i][colMap.role] ? data[i][colMap.role].toString().trim() : "Event Organizer";
      const rawsite = data[i][colMap.site] ? data[i][colMap.site].toString().trim().toUpperCase() : "Site1";
      const rawgroup = data[i][colMap.group] ? data[i][colMap.group].toString().trim() : "All";
      const rawPermLevel = data[i][colMap.permissionLevel] ? data[i][colMap.permissionLevel].toString().trim() : "";

      let allowedSites = [];
      if (rawsite === "ALL") {
        allowedSites = settings.sites.map(s => s.toUpperCase());
      } else if (settings.siteGroups[rawsite]) {
        allowedSites = settings.siteGroups[rawsite];
      } else {
        allowedSites = [rawsite];
      }

      let allowedEvents = [];
      if (rawPermLevel === "") {
        allowedEvents = [];
      } else if (rawPermLevel.toUpperCase() === "ALL") {
        allowedEvents = settings.events;
      } else {
        allowedEvents = rawPermLevel.split(",").map(e => e.trim()).filter(e => e);
      }

      return {
        role: rawRole,
        allowedSites: allowedSites,
        allowedgroup: rawgroup,
        permissionLevel: rawPermLevel,
        allowedEvents: allowedEvents
      };
    }
  }
  return null;
}

// =============================================================================
// SECTION 1: WEB APP (now passes app title and version)
// =============================================================================
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('sidebar');
  const activeUserEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
  const settings = getSettings();

  let userPerms = getRolePermissions(activeUserEmail);

  if (!userPerms) {
    template.config = {
      accessDenied: true,
      sites: [],
      events: [],
      userRole: "None",
      userSite: "",
      usergroup: "All",
      allowedSites: [],
      allowedgroup: "All",
      allowedEvents: [],
      userPermissionLevel: "",
      allowedSitesJson: "[]",
      allowedEventsJson: "[]",
      appTitle: settings.appTitle,
      appVersion: settings.appVersion
    };
    return template.evaluate()
      .setTitle("Event Check-In App")
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  template.config = {
    accessDenied: false,
    sites: settings.sites,
    events: userPerms.allowedEvents,
    userRole: userPerms.role,
    userSite: userPerms.allowedSites[0] || "",
    usergroup: userPerms.allowedgroup,
    allowedSites: userPerms.allowedSites,
    allowedgroup: userPerms.allowedgroup,
    allowedEvents: userPerms.allowedEvents,
    userPermissionLevel: userPerms.permissionLevel,
    allowedSitesJson: JSON.stringify(userPerms.allowedSites),
    allowedEventsJson: JSON.stringify(userPerms.allowedEvents),
    appTitle: settings.appTitle,
    appVersion: settings.appVersion
  };

  return template.evaluate()
    .setTitle("Event Check-In App")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =============================================================================
// SECTION 2: DYNAMIC EVENT SCHEDULE
// =============================================================================
function getEventScheduleBenchmark(eventType, site) {
  const settings = getSettings();
  const siteSchedules = settings.schedules[site];
  if (siteSchedules && siteSchedules[eventType]) return siteSchedules[eventType];
  for (let s in settings.schedules) {
    if (settings.schedules[s][eventType]) return settings.schedules[s][eventType];
  }
  return { hour: 11, minute: 0 };
}

// =============================================================================
// SECTION 3: AUTOMATIC ATTENDANCE SUBMISSION (with historical stats saving)
// =============================================================================
function submitAttendance(site, targetDateStr, eventType, personStatusMap) {
  if (!site || !site.trim()) return { success: false, message: "Missing site." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetDate = new Date(targetDateStr);
  targetDate.setHours(0,0,0,0);
  const targetYear = targetDate.getFullYear();
  let logSheet = ss.getSheetByName("Attendance Log_" + targetYear);
  if (!logSheet) logSheet = ss.getSheetByName("Attendance Log");
  if (!logSheet) return { success: false, message: "Attendance ledger sheet not found." };

  const now = new Date();
  const submitter = Session.getActiveUser().getEmail() || "Web App Portal";
  const today = new Date();
  today.setHours(0,0,0,0);
  const isToday = (targetDate.getTime() === today.getTime());

  const schedule = getEventScheduleBenchmark(eventType, site);
  const eventStartDeadline = new Date(now.getFullYear(), now.getMonth(), now.getDate(),
                                       schedule.hour, schedule.minute, 0);
  const twentyMinutesMillis = 20 * 60 * 1000;

  // Clear previous entries for this site/date/event
  const lastRow = logSheet.getLastRow();
  for (let i = lastRow; i >= 2; i--) {
    const rowDateRaw = logSheet.getRange(i, 2).getValue();
    if (!rowDateRaw) continue;
    const rowDate = new Date(rowDateRaw).setHours(0,0,0,0);
    const rowSite = logSheet.getRange(i, 3).getValue().toString().trim().toLowerCase();
    const rowEvent = logSheet.getRange(i, 4).getValue().toString().trim().toLowerCase();
    if (rowDate === targetDate.getTime() &&
        rowSite === site.toLowerCase() &&
        rowEvent === eventType.toLowerCase()) {
      logSheet.deleteRow(i);
    }
  }

  const newRows = [];
  const stats = { onTime: 0, late: 0, away: 0, absent: 0 };

  for (let personId in personStatusMap) {
    const status = personStatusMap[personId];
    let punctuality = "";
    if (status === "On-Time") {
      if (isToday && (now.getTime() - eventStartDeadline.getTime() > twentyMinutesMillis)) {
        punctuality = "L";  // auto-late for today
      } else {
        punctuality = "P";
      }
    } else if (status === "Late") {
      punctuality = "L";
    } else if (status === "Away") {
      punctuality = "A";
    } else {
      punctuality = "X"; // absent
    }

    if (punctuality === "P") stats.onTime++;
    else if (punctuality === "L") stats.late++;
    else if (punctuality === "A") stats.away++;
    else stats.absent++;

    if (personId) {
      newRows.push([now, targetDateStr, site, eventType, personId.toString().toUpperCase(), punctuality, submitter]);
    }
  }

  if (newRows.length > 0) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, newRows.length, 7).setValues(newRows);
  }

  // Save historical stats to "Daily Stats" sheet
  saveDailyStats(targetDateStr, site, eventType, stats);

  const label = `Saved: On-Time ${stats.onTime}, Late ${stats.late}, Away ${stats.away}, Absent ${stats.absent}`;
  return { success: true, message: label };
}

function saveDailyStats(dateStr, site, eventType, stats) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Daily Stats");
  if (!sheet) {
    sheet = ss.insertSheet("Daily Stats");
    sheet.appendRow(["Date", "Site", "Event", "On-Time", "Late", "Away", "Absent"]);
  }

  // Remove existing row for same date/site/event to avoid duplicates
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] && data[i][0].toString() === dateStr &&
        data[i][1] === site && data[i][2] === eventType) {
      sheet.deleteRow(i + 1);
    }
  }

  sheet.appendRow([dateStr, site, eventType, stats.onTime, stats.late, stats.away, stats.absent]);
}

// =============================================================================
// SECTION 4: DIRECTORY & PEOPLE DATA (new column layout)
// =============================================================================
function getAllPeopleInDatabase() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("People");
    if (!sheet) return {};
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return {};
    const output = {};
    // Column indices (0‑based):
    const idIdx = 0, firstIdx = 1, midIdx = 2, lastIdx = 3, phoneIdx = 4,
          demoIdx = 5, levelIdx = 6, famIdIdx = 7, siteIdx = 8, cgIdx = 9,
          statusIdx = 10, flagIdx = 11, notesIdx = 12;

    for (let i = 1; i < data.length; i++) {
      const id = data[i][idIdx] ? data[i][idIdx].toString().trim() : "";
      if (!id) continue;
      const first = data[i][firstIdx] ? data[i][firstIdx].toString().trim() : "";
      const middle = data[i][midIdx] ? data[i][midIdx].toString().trim() : "";
      const last = data[i][lastIdx] ? data[i][lastIdx].toString().trim() : "";
      const site = data[i][siteIdx] ? data[i][siteIdx].toString().trim() : "";
      const status = data[i][statusIdx] ? data[i][statusIdx].toString().trim() : "";
      const flag = data[i][flagIdx] ? data[i][flagIdx].toString().trim() : "";
      const notes = data[i][notesIdx] ? data[i][notesIdx].toString().trim() : "";
      const phone = data[i][phoneIdx] ? data[i][phoneIdx].toString().trim() : "";
      let fullName = first;
      if (middle && middle !== "-" && middle.toLowerCase() !== "nil") fullName += " " + middle;
      if (last) fullName += " " + last;
      fullName = fullName.replace(/\s+/g, ' ').trim();

      output[id] = {
        name: fullName,
        site: site,
        status: status,
        flag: flag,
        notes: notes,
        phone: phone,
        demographic: data[i][demoIdx] ? data[i][demoIdx].toString().trim() : "",
        level: data[i][levelIdx] ? data[i][levelIdx].toString().trim() : "",
        familyId: data[i][famIdIdx] ? data[i][famIdIdx].toString().trim() : "",
        group: data[i][cgIdx] ? data[i][cgIdx].toString().trim() : ""
      };
    }
    return output;
  } catch(e) {
    console.error("Error in getAllPeopleInDatabase:", e);
    return {};
  }
}

function getPeople(site, dateStr, eventType, role, group) {
  const output = [];
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("People");
    if (!sheet) return output;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return output;

    const idIdx = 0, firstIdx = 1, midIdx = 2, lastIdx = 3, phoneIdx = 4,
          demoIdx = 5, levelIdx = 6, famIdIdx = 7, siteIdx = 8, cgIdx = 9,
          statusIdx = 10, flagIdx = 11, notesIdx = 12;

    // Pre‑fetch attendance log for the given date/event (for pre‑populated checks and visitor counts)
    let loggedStatuses = {};
    let historicalCounts = {};
    if (dateStr && eventType) {
      loggedStatuses = getLoggedStatuses(site, dateStr, eventType);
      historicalCounts = getHistoricalVisitorAttendanceCounts(site, eventType);
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const personSite = row[siteIdx] ? row[siteIdx].toString().trim() : "";
      if (site !== "" && personSite.toLowerCase() !== site.toLowerCase()) continue;

      if (group && group !== "All") {
        const personCG = row[cgIdx] ? row[cgIdx].toString().trim() : "";
        if (personCG.toLowerCase() !== group.toLowerCase()) continue;
      }

      const pId = row[idIdx] ? row[idIdx].toString().trim() : "ID-" + i;
      const firstName = row[firstIdx] ? row[firstIdx].toString().trim() : "";
      const middleName = row[midIdx] ? row[midIdx].toString().trim() : "";
      const lastName = row[lastIdx] ? row[lastIdx].toString().trim() : "";
      let fullName = firstName;
      if (middleName && middleName !== "-" && middleName.toLowerCase() !== "nil") fullName += " " + middleName;
      if (lastName) fullName += " " + lastName;
      fullName = fullName.replace(/\s+/g, ' ').trim() || ("Unnamed Row " + (i+1));

      const rawDemo = row[demoIdx] ? row[demoIdx].toString().trim() : "";   // Family role or "university"
      const rawLevel = row[levelIdx] ? row[levelIdx].toString().trim() : ""; // University name & level
      const statusRole = row[statusIdx] ? row[statusIdx].toString().trim() : "Member"; // Member, Visitor, etc.
      const familyId = row[famIdIdx] ? row[famIdIdx].toString().trim() : "";
      const phoneRaw = row[phoneIdx] ? row[phoneIdx].toString().trim() : "";
      const missingPhone = (phoneRaw === "" || phoneRaw === "-" || phoneRaw.toLowerCase() === "nil");

      // Category grouping
      let masterCategory = "6. Others / Visitors";
      const lowerDemo = rawDemo.toLowerCase();
      if (familyId !== "") {
        masterCategory = "1. Families";
      } else if (lowerDemo.includes("double") || lowerDemo.includes("couple") || lowerDemo.includes("husband") || lowerDemo.includes("wife")) {
        masterCategory = "2. Doubles";
      } else if (lowerDemo.includes("single")) {
        masterCategory = (lowerDemo.includes("female") || lowerDemo.includes("woman") || lowerDemo.includes("girl"))
                         ? "4. Singles (Female)" : "3. Singles (Male)";
      } else if (lowerDemo.includes("student") || lowerDemo === "university") {
        masterCategory = "5. University Students";
      }

      // Determine initial status from log (pre‑populated check)
      let initialStatus = "Absent";
      let onTime = false, away = false;
      if (loggedStatuses[pId]) {
        const logStatus = loggedStatuses[pId];
        if (logStatus === "On-Time") { initialStatus = "On-Time"; onTime = true; }
        else if (logStatus === "Late") { initialStatus = "Late"; }
        else if (logStatus === "Away") { initialStatus = "Away"; away = true; }
      }

      // Visitor label – only for actual visitors, not all non‑members
      let visitorLabel = "";
      const isMember = (statusRole.toLowerCase() === "member");
      const isVisitorStatus = (statusRole.toLowerCase().includes("visitor") || statusRole.toLowerCase().includes("new"));
      const homeSite = row[siteIdx] ? row[siteIdx].toString().trim() : "";
      const visitCount = historicalCounts[pId] || 0;

      if (isMember) {
        // member visiting another site – show visiting label
        if (homeSite.toLowerCase() !== site.toLowerCase()) {
          visitorLabel = `${homeSite} Member-Visit #${visitCount + 1}`;
        }
      } else if (isVisitorStatus) {
        // Only people with a visitor status get the visit count
        visitorLabel = (visitCount === 0) ? "1st Visit" : `${visitCount + 1}th Visit`;
      }

      output.push({
        id: pId,
        name: fullName,
        masterCategory: masterCategory,
        familyId: familyId,
        demographic: rawDemo,
        level: rawLevel,
        role: statusRole,
        status: initialStatus,
        missingPhone: missingPhone,
        onTime: onTime,
        away: away,
        visitorLabel: visitorLabel,
        phone: phoneRaw,
        notes: row[notesIdx] ? row[notesIdx].toString().trim() : "",
        flag: row[flagIdx] ? row[flagIdx].toString().trim() : ""
      });
    }
    return output;
  } catch(e) {
    console.error("Error in getPeople:", e);
    return output;
  }
}

function getLoggedStatuses(site, dateStr, eventType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetDate = new Date(dateStr);
  targetDate.setHours(0,0,0,0);
  const targetYear = targetDate.getFullYear();
  let logSheet = ss.getSheetByName("Attendance Log_" + targetYear) || ss.getSheetByName("Attendance Log");
  if (!logSheet) return {};

  const data = logSheet.getDataRange().getValues();
  const statuses = {};
  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][1] ? new Date(data[i][1]) : null;
    if (!rowDate || rowDate.setHours(0,0,0,0) !== targetDate.getTime()) continue;
    const rowSite = data[i][2] ? data[i][2].toString().trim().toLowerCase() : "";
    const rowEvent = data[i][3] ? data[i][3].toString().trim().toLowerCase() : "";
    if (rowSite !== site.toLowerCase() || rowEvent !== eventType.toLowerCase()) continue;
    const personId = data[i][4] ? data[i][4].toString().trim().toUpperCase() : "";
    const punctuality = data[i][5] ? data[i][5].toString().trim() : "";
    let status = "Absent";
    if (punctuality === "P") status = "On-Time";
    else if (punctuality === "L") status = "Late";
    else if (punctuality === "A") status = "Away";
    if (personId) statuses[personId] = status;
  }
  return statuses;
}

function getHistoricalVisitorAttendanceCounts(site, eventType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Attendance Log");
  const trackingMap = {};
  if (!logSheet) return trackingMap;

  const data = logSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowSite = data[i][2] ? data[i][2].toString().trim().toLowerCase() : "";
    const rowEvent = data[i][3] ? data[i][3].toString().trim().toLowerCase() : "";
    const pid = data[i][4] ? data[i][4].toString().trim().toUpperCase() : "";
    if (rowSite === site.toLowerCase() && rowEvent === eventType.toLowerCase() && pid) {
      trackingMap[pid] = (trackingMap[pid] || 0) + 1;
    }
  }
  return trackingMap;
}

// =============================================================================
// SECTION 5: QUICK‑ADD NEW PERSON (adjusted for new columns, now accepts Level)
// =============================================================================
function quickAddPerson(site, first, middle, last, demographic, phone, level) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("People");
  if (!sheet) throw new Error('Sheet "People" not found.');
  // New order: ID, First, Middle, Last, Phone, Demographic, Level, FamilyID, site, CareGroup, Status, Flag, Notes
  sheet.appendRow([
    "", first, middle, last, phone, demographic, level || "", "", site, "", "Visitor", "", ""
  ]);
  renumberPersonIDs();
  return { success: true };
}

// =============================================================================
// SECTION 6: AUTO‑RENUMBER & SORT
// =============================================================================
function renumberPersonIDs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("People");
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const ids = [];
  for (let i = 2; i <= lastRow; i++) ids.push(["P" + String(i-1).padStart(5, "0")]);
  sheet.getRange(2, 1, ids.length, 1).setValues(ids);
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== "People") return;
  if (e.range.getColumn() === 1 && e.range.getRow() >= 2) return;
  renumberPersonIDs();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Database Tools')
    .addItem('Renumber Person IDs', 'renumberPersonIDs')
    .addToUi();
}

// =============================================================================
// SECTION 7: UPDATE (plus demographic update logging)
// =============================================================================
function updatePastoralCareStatus(personId, newFlag, newNoteText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("People");
  if (!sheet) return { success: false, message: "Sheet data missing" };

  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues(); // 13 columns
  const dateString = Utilities.formatDate(new Date(), "GMT+10", "dd/MM/yyyy");

  for (let i = 0; i < data.length; i++) {
    if (data[i][0].toString().trim().toUpperCase() === personId.toUpperCase()) {
      const rowPosition = i + 2;
      let existingNotes = data[i][12] ? data[i][12].toString().trim() : ""; // Notes column (M)

      sheet.getRange(rowPosition, 12).setValue(newFlag); // Flag column (L)

      let appendString = "";
      if (newNoteText && newNoteText.trim() !== "") {
        appendString = "[" + dateString + "]: " + newNoteText.trim();
      } else {
        appendString = "[" + dateString + "]: Care Status adjusted to '" + (newFlag || "None / Cleared") + "'";
      }

      const finalNotesBlock = existingNotes ? appendString + "\n" + existingNotes : appendString;
      sheet.getRange(rowPosition, 13).setValue(finalNotesBlock);
      SpreadsheetApp.flush();
      return { success: true, updatedFlag: newFlag, updatedNotes: finalNotesBlock };
    }
  }
  return { success: false };
}

function updateVisitorPhone(personId, newPhoneNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("People");
  if (!sheet) return { success: false };
  const lastRow = sheet.getLastRow();
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0].toString().trim().toUpperCase() === personId.toUpperCase()) {
      sheet.getRange(i + 2, 5).setValue(newPhoneNumber); // Phone column (E)
      SpreadsheetApp.flush();
      return { success: true };
    }
  }
  return { success: false };
}

/**
 * NEW: Update demographic and log the change (e.g., from Child to University)
 */
function updateDemographic(personId, newDemographic) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("People");
  if (!sheet) return { success: false, message: "Sheet data missing" };

  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const dateString = Utilities.formatDate(new Date(), "GMT+10", "dd/MM/yyyy");

  for (let i = 0; i < data.length; i++) {
    if (data[i][0].toString().trim().toUpperCase() === personId.toUpperCase()) {
      const rowPosition = i + 2;
      const oldDemographic = data[i][5] ? data[i][5].toString().trim() : ""; // Demographic column (F)
      
      // Update the demographic column
      sheet.getRange(rowPosition, 6).setValue(newDemographic); // column F (1-based = 6)

      // If changing to "University", log it in notes
      if (newDemographic.toLowerCase() === "university" && oldDemographic.toLowerCase() !== "university") {
        let existingNotes = data[i][12] ? data[i][12].toString().trim() : "";
        const noteEntry = `[${dateString}]: Now attending university.`;
        const finalNotes = existingNotes ? noteEntry + "\n" + existingNotes : noteEntry;
        sheet.getRange(rowPosition, 13).setValue(finalNotes);
        SpreadsheetApp.flush();
        return { success: true, message: "Demographic updated and university note logged." };
      }

      SpreadsheetApp.flush();
      return { success: true, message: "Demographic updated." };
    }
  }
  return { success: false, message: "Person not found." };
}

// =============================================================================
// SECTION 8: PERMISSION‑AWARE STATS
// =============================================================================
function getFilteredStats(dateStr, eventType, userRole, allowedSites) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetDate = new Date(dateStr);
  targetDate.setHours(0,0,0,0);
  const targetYear = targetDate.getFullYear();

  const stats = {};
  allowedSites.forEach(site => {
    stats[site] = { name: site, onTime: 0, late: 0, away: 0, absent: 0 };
  });

  let logSheet = ss.getSheetByName("Attendance Log_" + targetYear) ||
                 ss.getSheetByName("Attendance Log");
  if (logSheet) {
    const logData = logSheet.getDataRange().getValues();
    for (let i = 1; i < logData.length; i++) {
      const rowDate = logData[i][1] ? new Date(logData[i][1]) : null;
      if (!rowDate || rowDate.setHours(0,0,0,0) !== targetDate.getTime()) continue;
      const rowEvent = logData[i][3] ? logData[i][3].toString().trim().toLowerCase() : "";
      const rowSite = logData[i][2] ? logData[i][2].toString().trim() : "";
      const punctuality = logData[i][5] ? logData[i][5].toString().trim().toUpperCase() : "";
      if (rowEvent !== eventType.toLowerCase()) continue;
      if (!allowedSites.includes(rowSite)) continue;
      if (punctuality === "P") stats[rowSite].onTime++;
      else if (punctuality === "L") stats[rowSite].late++;
      else if (punctuality === "A") stats[rowSite].away++;
      else stats[rowSite].absent++;
    }
  }

  const peopleSheet = ss.getSheetByName("People");
  if (peopleSheet) {
    const data = peopleSheet.getDataRange().getValues();
    allowedSites.forEach(site => {
      let expected = 0;
      for (let i = 1; i < data.length; i++) {
        if (data[i][8] && data[i][8].toString().trim() === site) expected++; // Site column (I)
      }
      stats[site].absent = expected;
    });
    for (let site in stats) {
      const s = stats[site];
      s.absent = Math.max(0, s.absent - (s.onTime + s.late + s.away));
    }
  }

  const global = { onTime: 0, late: 0, away: 0, absent: 0 };
  for (let site in stats) {
    global.onTime += stats[site].onTime;
    global.late += stats[site].late;
    global.away += stats[site].away;
    global.absent += stats[site].absent;
  }
  return { global: global, sites: stats };
}
