const SHEET_ID = "1nxxwp3XG9cNWcGcnCkwp5OfBQxeGl-qqY4ikW4WzLQg";

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('WorkforcePay System')
    .setFaviconUrl('https://i.ibb.co/60M2J7YZ/image.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
// Add this to your code.gs
function checkLogin(username, password) {
  // Replace these with your desired credentials or a sheet lookup
  const ADMIN_USER = "admin";
  const ADMIN_PASS = "123"; 
  
  if (username === ADMIN_USER && password === ADMIN_PASS) {
    return true;
  }
  return false;
}

// DASHBOARD
function getEmployeeList() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("SETTINGS");
  const data = sheet.getDataRange().getValues();
  
  let employees = [];
  // Start from row 2 (index 1) to skip header
  for (let i = 1; i < data.length; i++) {
    if (data[i][1]) { // Full Name is in Column B
      employees.push(data[i][1].toString().trim()); 
    }
  }
  return [...new Set(employees)].sort(); 
}
function getDashboardStats() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  // Get Today's Date at Midnight for accurate comparison
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");

  // 1. Get Directory Data
  const settingsSheet = ss.getSheetByName("SETTINGS");
  const settingsData = settingsSheet.getDataRange().getValues();
  let directory = [];
  for(let i = 1; i < settingsData.length; i++){
    if(settingsData[i][1]) {
      directory.push({
        id: settingsData[i][0],
        name: settingsData[i][1].toString().trim(),
        dept: settingsData[i][2],
        pos: settingsData[i][3]
      });
    }
  }

  // 2. Get Logs for Present/Leave
  const logSheet = ss.getSheetByName("DTR LOGS");
  const logs = logSheet.getDataRange().getValues();
  let presentNames = new Set();
  let onLeaveNames = new Set();

  for (let i = 1; i < logs.length; i++) {
    if (!logs[i][0]) continue;
    
    // FIX 1: Robust Date Comparison
    let logDateValue = new Date(logs[i][0]);
    let logDateStr = Utilities.formatDate(logDateValue, "GMT+8", "yyyy-MM-dd");
    
    let name = logs[i][2] ? logs[i][2].toString().trim() : "";
    let remarks = logs[i][5] ? logs[i][5].toString().toUpperCase() : "";

    if (logDateStr === todayStr) {
      if (remarks.includes("LEAVE")) {
        onLeaveNames.add(name);
      } else {
        presentNames.add(name);
      }
    }
  }

  // 3. Match On-Leave names with Detailed Monitoring Data
const leaveSheet = ss.getSheetByName("LEAVE MONITORING");
  const leaveData = leaveSheet.getDataRange().getValues();
  let activeLeaves = [];
  
  onLeaveNames.forEach(name => {
    // FIX: Added checks to ensure row and row[2] exist before calling toString()
    let leaveRecord = leaveData.slice().reverse().find(row => {
      return row && row[2] && row[2].toString().trim() === name && 
             (row[17] && row[17].toString().toUpperCase() === "TRANSFERRED");
    });

    activeLeaves.push({
      name: name,
      type: leaveRecord ? leaveRecord[4] : "Leave", 
      // Safely format dates, fallback to "Today" if missing
      start: (leaveRecord && leaveRecord[5] instanceof Date) 
             ? Utilities.formatDate(leaveRecord[5], "GMT+8", "MMM dd") : "Today", 
      end: (leaveRecord && leaveRecord[6] instanceof Date) 
             ? Utilities.formatDate(leaveRecord[6], "GMT+8", "MMM dd") : "Today"
    });
  });

  return {
    total: { count: directory.length, list: directory },
    present: { count: presentNames.size, list: Array.from(presentNames) },
    onLeave: { count: activeLeaves.length, list: activeLeaves },
    pending: { count: getPendingLeaves().length, list: getPendingLeaves() }
  };
}

function getSettingsData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("SETTINGS");
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  // Binago: Mula 20, ginawa nating 21 para makuha ang Column U (Exemption)
  return sheet.getRange(2, 1, lastRow - 1, 21).getDisplayValues(); 
}

function saveEmployee(...vals) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("SETTINGS");
  const data = sheet.getDataRange().getValues();
  const nameToFind = vals[1] ? vals[1].trim() : ""; // Full Name (Col B)
  let found = false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().trim() === nameToFind) {
      // FIX: Binago ang column count mula 20 tungong 21
      sheet.getRange(i + 1, 1, 1, 21).setValues([vals]);
      found = true;
      break;
    }
  }
  
  if (!found) {
    // Siguraduhin na ang array ay sakto sa columns ng sheet para iwas error sa appendRow
    sheet.appendRow(vals);
  }
  return "Success";
}

function getHRReport(filterStart, filterEnd, employeeFilter) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const logSheet = ss.getSheetByName("DTR LOGS");
  const settingsSheet = ss.getSheetByName("SETTINGS");
  
  const logs = logSheet.getDataRange().getValues();
  const settingsData = settingsSheet.getDataRange().getValues();
  
  // 1. I-normalize ang filter dates (Alisin ang oras para date lang ang labanan)
  const startDate = new Date(filterStart);
  startDate.setHours(0, 0, 0, 0);
  const endDate = new Date(filterEnd);
  endDate.setHours(23, 59, 59, 999);
  
  const isSecondCutoff = endDate.getDate() > 15;

  // Map Settings & identify Exempted Employees
  let config = {};
  for (let i = 1; i < settingsData.length; i++) {
    let r = settingsData[i];
    if (r[1]) { 
      let fullName = r[1].toString().trim();
      config[fullName] = { 
        id: r[0],
        rate: parseFloat(r[8]) || 0,
        hono: parseFloat(r[13]) || 0,
        sss: isSecondCutoff ? (parseFloat(r[16]) || 0) : 0,
        philhealth: isSecondCutoff ? (parseFloat(r[17]) || 0) : 0,
        pagibig: isSecondCutoff ? (parseFloat(r[18]) || 0) : 0,
        wTax: isSecondCutoff ? (parseFloat(r[19]) || 0) : 0,
        isExempted: r[20] && r[20].toString().toUpperCase() === "EXEMPTED" 
      };
    }
  }

  // 2. Generate Date List (Ito ang magsisilbing basehan ng rows sa table)
  let dateList = [];
  let curr = new Date(startDate);
  while (curr <= endDate) {
    dateList.push(Utilities.formatDate(new Date(curr), "GMT+8", "yyyy-MM-dd"));
    curr.setDate(curr.getDate() + 1);
  }

  let reportMap = {};
  let empSummaries = {};

  // 3. Process Logs (Siguraduhin na pasok sa range)
  for (let i = 1; i < logs.length; i++) {
    let row = logs[i];
    if (!row[0]) continue;
    
    let timestamp = new Date(row[0]);
    if (isNaN(timestamp.getTime())) continue; // Skip kung invalid date

    let name = row[2] ? row[2].toString().trim() : "";
    let status = row[3] ? row[3].toString().trim() : "";
    
    if (!name || (employeeFilter && employeeFilter !== "ALL" && name !== employeeFilter)) continue;
    
    // I-check kung ang log ay nasa loob ng filter range
    if (timestamp < startDate || timestamp > endDate) continue;

    let dateKey = Utilities.formatDate(timestamp, "GMT+8", "yyyy-MM-dd");
    let key = name + "_" + dateKey;
    
    if (!reportMap[key]) {
      reportMap[key] = { name: name, date: dateKey, in: null, out: null, isLeave: false, leaveType: "", remarks: "" };
    }
    
    if (status === "IN") reportMap[key].in = timestamp;
    if (status === "OUT") reportMap[key].out = timestamp;
    if (row[5] && row[5].toString().toUpperCase().includes("LEAVE")) {
      reportMap[key].isLeave = true;
      reportMap[key].leaveType = row[5].toString().replace(/LEAVE:\s*/i, "").trim();
    }
    if (row[5]) reportMap[key].remarks = row[5].toString();
  }

  let finalReport = [];

  // 4. Loop through Configured Employees x Date List
  for (let name in config) {
    if (employeeFilter && employeeFilter !== "ALL" && name !== employeeFilter) continue;

    let userConf = config[name];

    dateList.forEach(dateKey => {
      let item = reportMap[name + "_" + dateKey] || { name: name, date: dateKey, in: null, out: null, isLeave: false, remarks: "" };
      
      let currentDayObj = new Date(dateKey);
      let dayOfWeek = currentDayObj.getDay(); 

      let lateMin = 0, otMin = 0, ded = 0, payOT = 0, leavePayValue = 0;
      let dailyGross = userConf.rate * 8; 
      let isExempted = userConf.isExempted;

      let hasIn = item.in !== null;
      let hasOut = item.out !== null;
      let hasAnyLog = hasIn || hasOut;
      
      let isSaturday = (dayOfWeek === 6);
      let shouldAutoPay = isExempted && !hasAnyLog && !item.isLeave && !isSaturday;

      if (item.isLeave) {
        leavePayValue = dailyGross;
      } else if (hasAnyLog) {
        if (hasIn && !isExempted) {
          let lateStartLimit = new Date(item.in);
          lateStartLimit.setHours(9, 1, 0, 0); 
          if (item.in >= lateStartLimit) {
             let officialSchedIn = new Date(item.in);
             officialSchedIn.setHours(8, 30, 0, 0);
             lateMin = (item.in - officialSchedIn) / (1000 * 60);
             ded = (userConf.rate / 60) * lateMin;
          }
        }
        if (hasOut) {
          let otStartLimit = new Date(item.out);
          otStartLimit.setHours(18, 0, 0, 0); 
          if (item.out >= otStartLimit) {
            otMin = (item.out - otStartLimit) / (1000 * 60);
            payOT = (userConf.rate / 60) * otMin;
          }
        }
      } else if (!shouldAutoPay) {
        return; // Skip day kung walang log at hindi auto-pay
      }

      let isPaidDay = item.isLeave || (hasIn && hasOut) || shouldAutoPay;
      let netDaily = isPaidDay ? (dailyGross + payOT - ded) : 0;

      if (!empSummaries[name]) {
        empSummaries[name] = { 
          name: name, days: 0, totalDed: 0, totalOT: 0, gross: 0, 
          sss: userConf.sss, philhealth: userConf.philhealth, 
          pagibig: userConf.pagibig, wTax: userConf.wTax, hono: userConf.hono,
          leaveDays: 0, totalLeavePay: 0, leaveType: ""
        };
      }
      
      let s = empSummaries[name];
      if (item.isLeave) {
        s.leaveDays += 1;
        s.totalLeavePay += dailyGross;
        s.leaveType = item.leaveType;
      } else if (hasAnyLog || shouldAutoPay) {
        s.days += 1;
        s.gross += dailyGross;
      }
      s.totalDed += ded;
      s.totalOT += payOT;

      finalReport.push({
        name: name, 
        date: dateKey,
        timestamp: hasIn ? Utilities.formatDate(item.in, "GMT+8", "yyyy-MM-dd HH:mm:ss") : (hasOut ? Utilities.formatDate(item.out, "GMT+8", "yyyy-MM-dd HH:mm:ss") : ""),
        in: item.isLeave ? item.leaveType : (hasIn ? Utilities.formatDate(item.in, "GMT+8", "hh:mm a") : (shouldAutoPay ? "EXEMPTED" : "MISSING")),
        out: item.isLeave ? "" : (hasOut ? Utilities.formatDate(item.out, "GMT+8", "hh:mm a") : (shouldAutoPay ? "AUTO-PAY" : "MISSING")),
        status: item.isLeave ? "LEAVE" : (shouldAutoPay ? "FIXED" : "WORK"),
        dailyNet: netDaily.toFixed(2),
        deduction: ded.toFixed(2), 
        otPay: payOT.toFixed(2),
        leavePay: leavePayValue.toFixed(2),
        remarks: item.remarks
      });
    });
  }

  let empList = Object.values(empSummaries).map(s => {
    let userConf = config[s.name] || {};
    s.id = userConf.id || "N/A";
    s.rate = userConf.rate || 0;
    s.totalEarnings = s.gross + s.totalOT + s.hono + s.totalLeavePay;
    s.totalFixedDeductions = s.pagibig + s.philhealth + s.sss + s.wTax + s.totalDed;
    s.finalNetPay = s.totalEarnings - s.totalFixedDeductions;
    return s;
  });

  return { 
    details: finalReport.sort((a,b) => b.date.localeCompare(a.date)), 
    empList: empList 
  };
}


// code.gs - ADD THIS FUNCTION
/**
 * Saves or updates a manual log entry.
 * It looks for an existing IN or OUT for that person on that day.
 */
function saveManualLog(name, dateStr, timeIn, timeOut, remark) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("DTR LOGS");
  const data = sheet.getDataRange().getValues();
  
  // 1. Identify existing rows for this employee on this specific date
  const existingRows = [];
  for (let i = 1; i < data.length; i++) {
    // Format the date from the sheet to match YYYY-MM-DD
    let rowDate = Utilities.formatDate(new Date(data[i][0]), "GMT+8", "yyyy-MM-dd");
    if (rowDate === dateStr && data[i][2] === name) {
      existingRows.push({ index: i + 1, status: data[i][3] });
    }
  }

  // 2. Process Time In
  if (timeIn) {
    const inRow = existingRows.find(r => r.status === "IN");
    const timestampIn = dateStr + " " + timeIn + ":00";
    if (inRow) {
      // Update existing
      sheet.getRange(inRow.index, 1).setValue(timestampIn);
      sheet.getRange(inRow.index, 6).setValue(remark);
      sheet.getRange(inRow.index, 4).setValue("IN"); // Ensure status is IN
    } else {
      // Create new row: [Timestamp, "", Name, Status, DateOnly, Remark]
      sheet.appendRow([timestampIn, "", name, "IN", dateStr, remark]);
    }
  }

  // 3. Process Time Out
  if (timeOut) {
    const outRow = existingRows.find(r => r.status === "OUT");
    const timestampOut = dateStr + " " + timeOut + ":00";
    if (outRow) {
      // Update existing
      sheet.getRange(outRow.index, 1).setValue(timestampOut);
      sheet.getRange(outRow.index, 6).setValue(remark);
      sheet.getRange(outRow.index, 4).setValue("OUT");
    } else {
      sheet.appendRow([timestampOut, "", name, "OUT", dateStr, remark]);
    }
  }

  return "Success";
}


// LEAVE 
function transferLeaveToDTR(rowNum, name, startDate, endDate, leaveType) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const logSheet = ss.getSheetByName("DTR LOGS");
  const leaveSheet = ss.getSheetByName("LEAVE MONITORING");
  
  let start = new Date(startDate);
  let end = new Date(endDate);
  
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    let dateStr = Utilities.formatDate(d, "GMT+8", "yyyy-MM-dd");
    let timestampIn = dateStr + " 08:30:00";
    let timestampOut = dateStr + " 17:30:00";
    logSheet.appendRow([timestampIn, "", name, "IN", dateStr, "LEAVE: " + leaveType]);
    logSheet.appendRow([timestampOut, "", name, "OUT", dateStr, "LEAVE: " + leaveType]);
  }
  
  leaveSheet.getRange(rowNum, 18).setValue("TRANSFERRED");
  return "Success";
}

function getPendingLeaves() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("LEAVE MONITORING");
  const data = sheet.getDataRange().getValues();
  
  let pending = [];
  for (let i = 1; i < data.length; i++) {
    // Column indices based on your image_e9771d.png:
    // J=9 (Imm Head), K=10 (Op Head), R=17 (Status)
    let immHead = data[i][9] ? data[i][9].toString().toUpperCase() : "";
    let opHead = data[i][10] ? data[i][10].toString().toUpperCase() : "";
    let transferStatus = data[i][17] ? data[i][17].toString().toUpperCase() : "";

    if (immHead === "APPROVED" && opHead === "APPROVED" && transferStatus !== "TRANSFERRED") {
      pending.push({
        row: i + 1,
        name: data[i][2], // Column C
        type: data[i][4], // Column E
        start: Utilities.formatDate(new Date(data[i][5]), "GMT+8", "MMM dd, yyyy"), 
        end: Utilities.formatDate(new Date(data[i][6]), "GMT+8", "MMM dd, yyyy"), 
        days: data[i][7], // Column H
        reason: data[i][8] || "No reason provided", // Column I
        dept: data[i][3] // Column D
      });
    }
  }
  return pending;
}


// ATTENDANCE
function deleteEmployeeFromSheet(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("SETTINGS");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === name) {
      sheet.deleteRow(i + 1);
      return "Success";
    }
  }
  return "Employee not found";
}


function updateDTRRow(originalTimestamp, employeeName, newStatus, newTime, newRemark) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("DTR LOGS");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    let rowDateStr = Utilities.formatDate(new Date(data[i][0]), "GMT+8", "yyyy-MM-dd HH:mm:ss");
    if (rowDateStr === originalTimestamp && data[i][2] === employeeName) {
      
      const targetRow = i + 1;
      // Kunin natin yung date part lang nung original para hindi mabura ang petsa
      const datePart = rowDateStr.split(" ")[0]; 
      const updatedTimestamp = datePart + " " + newTime + ":00";

      sheet.getRange(targetRow, 1).setValue(updatedTimestamp); 
      sheet.getRange(targetRow, 4).setValue(newStatus); 
      sheet.getRange(targetRow, 6).setValue(newRemark); 
      return "Success";
    }
  }
  return "Error: Row not found";
}

function getRawLogs(dateStart, dateEnd) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("DTR LOGS");
  const data = sheet.getDataRange().getValues();
  
  // Normalize dates to exclude time for strict date filtering
  const start = new Date(dateStart).setHours(0,0,0,0);
  const end = new Date(dateEnd).setHours(23,59,59,999);

  return data.slice(1).filter(row => {
    if (!row[0]) return false;
    const logDate = new Date(row[0]).getTime();
    return logDate >= start && logDate <= end;
  }).map(row => ({
    timestamp: Utilities.formatDate(new Date(row[0]), "GMT+8", "yyyy-MM-dd HH:mm:ss"),
    displayTime: Utilities.formatDate(new Date(row[0]), "GMT+8", "hh:mm a"),
    dateOnly: row[4], 
    name: row[2],
    status: row[3],
    remarks: row[5] || "" // Fix for the "undefined" issue
  }));
}

// Function para sa Bulk Adjustment (Lahat o maraming empleyado)
function saveBulkLogs(employeeNames, dateStr, timeIn, timeOut, remark) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("DTR LOGS");
  const data = sheet.getDataRange().getValues();
  let addedCount = 0;

  employeeNames.forEach(name => {
    // Kunin lahat ng logs ng specific employee sa araw na yan
    const existingLogs = data.filter(row => {
      const rowDate = Utilities.formatDate(new Date(row[0]), "GMT+8", "yyyy-MM-dd");
      return rowDate === dateStr && row[2] === name;
    });

    const hasIn = existingLogs.some(row => row[3] === "IN");
    const hasOut = existingLogs.some(row => row[3] === "OUT");

    // Mag-a-append lang kung ENABLED ang timeIn at WALA pang existing IN
    if (timeIn && !hasIn) {
      sheet.appendRow([dateStr + " " + timeIn + ":00", "", name, "IN", dateStr, remark]);
      addedCount++;
    }

    // Mag-a-append lang kung ENABLED ang timeOut at WALA pang existing OUT
    if (timeOut && !hasOut) {
      sheet.appendRow([dateStr + " " + timeOut + ":00", "", name, "OUT", dateStr, remark]);
      addedCount++;
    }
  });

  return `Bulk Fix Complete: Added ${addedCount} missing logs. Existing logs were preserved.`;
}

function deleteLogRow(timestamp, employeeName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("DTR LOGS");
  const data = sheet.getDataRange().getValues();
  
  // Use a standard for loop with 'let' to avoid scoping issues
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    
    // Create a string that matches the format sent by the JS
    let rowDateStr = Utilities.formatDate(new Date(data[i][0]), "GMT+8", "yyyy-MM-dd HH:mm:ss");
    
    if (rowDateStr === timestamp && data[i][2] === employeeName) {
      sheet.deleteRow(i + 1);
      return "Success";
    }
  }
  return "Error: Row not found for " + employeeName + " at " + timestamp;
}