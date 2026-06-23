 COMPANY MANAGEMENT SYSTEM - Google Apps Script Backend
//  All data stored in Google Sheets
// ============================================================
var SPREADSHEET_ID = ''; // Leave empty to use active spreadsheet
//  COMPANY MANAGEMENT SYSTEM - Google Apps Script Backend v2
//  Fixed: Date handling, attendance save/update, dashboard sync
// ============================================================
var SS_ID = ''; // Leave empty — script uses the bound spreadsheet
function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
  return SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActiveSpreadsheet();
}
// ---- DATE UTILITIES ----
function toDateStr(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(val).trim();
  if (s.match(/^\d{4}-\d{2}-\d{2}/)) return s.substring(0, 10);
  try {
    var d = new Date(s);
    if (!isNaN(d)) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch(e) {}
  return s;
}
function toTimeStr(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
  return String(val).trim();
}
function todayStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
// ---- SERVE HTML ----
-2
+2
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
// ---- INITIALIZE ALL SHEETS ----
// ---- SHEET INITIALIZATION ----
function initSheets() {
  var ss = getSpreadsheet();
  var sheets = {
  var defs = {
    'Employees':      ['EmpID','Name','FatherName','CNIC','Phone','Address','Department','Designation','JoiningDate','BasicSalary','Status'],
    'Attendance':     ['AttID','EmpID','EmpName','Date','CheckIn','CheckOut','Status'],
    'Salaries':       ['SalID','EmpID','EmpName','Month','Year','BasicSalary','Allowances','Deductions','NetSalary','Paid','PaidDate','Remarks'],
-51
+112
    'JournalEntries': ['JrnID','Date','AccountDebit','AccountCredit','Amount','Description','Reference'],
    'Accounts':       ['AccID','AccountName','Type','OpeningBalance'],
    'Inventory':      ['InvtID','ItemName','Category','StockIn','StockOut','Remaining','CostPrice','SellingPrice','Unit'],
    'Users':          ['UserID','Username','Password','Role','Name','Status']
    'Users':          ['UserID','Username','PasswordHash','Role','Name','Status','LastLogin']
  };
  for (var name in sheets) {
  for (var name in defs) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(sheets[name]);
      // Style header row
      sheet.getRange(1, 1, 1, sheets[name].length)
        .setBackground('#1a73e8')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      sheet.appendRow(defs[name]);
      var hdr = sheet.getRange(1, 1, 1, defs[name].length);
      hdr.setBackground('#1a73e8').setFontColor('#fff').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  // Seed default admin user if Users sheet is empty
  // Default admin user
  var usersSheet = ss.getSheetByName('Users');
  if (usersSheet.getLastRow() <= 1) {
    usersSheet.appendRow(['U001','admin','admin123','Admin','Administrator','Active']);
    usersSheet.appendRow(['U-001','admin','admin123','Admin','Administrator','Active','']);
  }
  // Seed default accounts if empty
  // Default accounts
  var accSheet = ss.getSheetByName('Accounts');
  if (accSheet.getLastRow() <= 1) {
    var defaultAccounts = [
      ['A001','Cash','Asset',0],
      ['A002','Bank','Asset',0],
      ['A003','Sales Income','Income',0],
      ['A004','Salary Expense','Expense',0],
      ['A005','Rent Expense','Expense',0],
      ['A006','Electricity Expense','Expense',0],
      ['A007','Transport Expense','Expense',0],
      ['A008','Other Expense','Expense',0],
      ['A009','Accounts Receivable','Asset',0],
      ['A010','Retained Earnings','Equity',0]
    ];
    defaultAccounts.forEach(function(row) { accSheet.appendRow(row); });
    [['A-001','Cash','Asset',0],['A-002','Bank','Asset',0],['A-003','Sales Income','Income',0],
     ['A-004','Salary Expense','Expense',0],['A-005','Rent Expense','Expense',0],
     ['A-006','Electricity Expense','Expense',0],['A-007','Transport Expense','Expense',0],
     ['A-008','Other Expense','Expense',0],['A-009','Accounts Receivable','Asset',0]
    ].forEach(function(r){ accSheet.appendRow(r); });
  }
}
// ---- UTILITY ----
function generateID(prefix, sheet) {
function nextID(sheet, prefix) {
  var s = getSpreadsheet().getSheetByName(sheet);
  var last = s.getLastRow(); // includes header
  var num = last; // so first real row => num=1
  return prefix + '-' + String(num).padStart(3, '0');
}
function getSheetData(name) {
  var ss = getSpreadsheet();
  var s = ss.getSheetByName(sheet);
  var last = s.getLastRow();
  return prefix + String(last).padStart(4, '0');
}
function getSheetData(sheetName) {
  var s = ss.getSheetByName(name);
  if (!s || s.getLastRow() <= 1) return [];
  var vals = s.getDataRange().getValues();
  var headers = vals[0];
  return vals.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      var v = row[i];
      // Convert Date objects
      if (v instanceof Date) {
        if (headers[i] === 'Date' || headers[i] === 'JoiningDate' || headers[i] === 'PaidDate') {
          v = toDateStr(v);
        } else if (headers[i] === 'CheckIn' || headers[i] === 'CheckOut') {
          v = toTimeStr(v);
        } else if (headers[i] === 'LastLogin') {
          v = toDateStr(v);
        }
      }
      obj[h] = v;
    });
    return obj;
  });
}
function findRowByID(sheetName, id) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  return rows.map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}
function getRowIndexById(sheetName, idCol, idValue) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idValue)) return i + 1;
  var s = ss.getSheetByName(sheetName);
  var vals = s.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(id)) return i + 1; // 1-based row
  }
  return -1;
}
function today() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
// ============================================================
//  AUTH
// ============================================================
function loginUser(username, password) {
  try {
    var users = getSheetData('Users');
    var u = users.find(function(u) {
      return String(u.Username).toLowerCase() === String(username).toLowerCase() &&
             String(u.PasswordHash) === String(password) &&
             u.Status === 'Active';
    });
    if (u) {
      // Update last login
      var row = findRowByID('Users', u.UserID);
      if (row > 0) {
        var s = getSpreadsheet().getSheetByName('Users');
        s.getRange(row, 7).setValue(todayStr());
      }
      return { success: true, user: { UserID: u.UserID, Username: u.Username, Role: u.Role, Name: u.Name } };
    }
    return { success: false, error: 'Invalid username or password' };
  } catch(e) { return { success: false, error: e.message }; }
}
function getUsers() {
  return getSheetData('Users').map(function(u) {
    return { UserID: u.UserID, Username: u.Username, Role: u.Role, Name: u.Name, Status: u.Status, LastLogin: u.LastLogin };
  });
}
function addUser(data) {
  try {
    var users = getSheetData('Users');
    if (users.find(function(u){ return String(u.Username).toLowerCase() === String(data.Username).toLowerCase(); })) {
      return { success: false, error: 'Username already exists' };
    }
    var id = nextID('Users', 'U');
    getSpreadsheet().getSheetByName('Users').appendRow([id, data.Username, data.Password, data.Role, data.Name, 'Active', '']);
    return { success: true, id: id };
  } catch(e) { return { success: false, error: e.message }; }
}
function updateUser(data) {
  try {
    var s = getSpreadsheet().getSheetByName('Users');
    var row = findRowByID('Users', data.UserID);
    if (row < 0) return { success: false, error: 'User not found' };
    s.getRange(row, 2, 1, 5).setValues([[data.Username, data.Password || s.getRange(row, 3).getValue(), data.Role, data.Name, data.Status]]);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}
function changePassword(userID, oldPass, newPass) {
  try {
    var s = getSpreadsheet().getSheetByName('Users');
    var row = findRowByID('Users', userID);
    if (row < 0) return { success: false, error: 'User not found' };
    var current = s.getRange(row, 3).getValue();
    if (String(current) !== String(oldPass)) return { success: false, error: 'Old password incorrect' };
    s.getRange(row, 3).setValue(newPass);
    return { success: true, message: 'Password changed successfully' };
  } catch(e) { return { success: false, error: e.message }; }
}
// ============================================================
-35
+14
function getDashboardStats() {
  try {
    var employees = getSheetData('Employees');
    var salaries = getSheetData('Salaries');
    var invoices = getSheetData('Invoices');
    var expenses = getSheetData('Expenses');
    var attendance = getSheetData('Attendance');
    var activeEmp = employees.filter(function(e) { return e.Status === 'Active'; }).length;
    var todayStr = today();
    var todayAtt = attendance.filter(function(a) {
      return String(a.Date).substring(0,10) === todayStr;
    });
    var presentToday = todayAtt.filter(function(a) { return a.Status === 'Present'; }).length;
    var absentToday = todayAtt.filter(function(a) { return a.Status === 'Absent'; }).length;
    var salaries  = getSheetData('Salaries');
    var invoices  = getSheetData('Invoices');
    var expenses  = getSheetData('Expenses');
    var attendance= getSheetData('Attendance');
    var today = todayStr();
    var now = new Date();
    var thisMonth = now.getMonth() + 1;
    var thisYear = now.getFullYear();
    var monthlySalary = salaries
      .filter(function(s) { return Number(s.Month) === thisMonth && Number(s.Year) === thisYear; })
      .reduce(function(sum, s) { return sum + Number(s.NetSalary || 0); }, 0);
    var pendingSalary = salaries
      .filter(function(s) { return s.Paid !== 'Yes'; })
      .reduce(function(sum, s) { return sum + Number(s.NetSalary || 0); }, 0);
    var totalIncome = invoices.reduce(function(sum, inv) { return sum + Number(inv.PaidAmount || 0); }, 0);
    var pendingInvoice = invoices
      .filter(function(inv) { return inv.Status === 'Unpaid' || inv.Status === 'Partial'; })
      .reduce(function(sum, inv) { return sum + (Number(inv.TotalAmount || 0) - Number(inv.PaidAmount || 0)); }, 0);
    var totalExpenses = expenses.reduce(function(sum, e) { return sum + Number(e.Amount || 0); }, 0);
    var netProfit = totalIncome - totalExpenses;
    var recentInvoices = invoices.slice(-5).reverse();
    var recentExpenses = expenses.slice(-5).reverse();
    var thisYear  = now.getFullYear();
    var activeEmp = employees.filter(function(e){ return e.Status === 'Active'; }).length;
    var todayAtt = attendance.filter(function(a){ return toDateStr(a.Date) === today; });
    var presentToday = todayAtt.filter(function(a){ return a.Status === 'Present'; }).length;
    var absentToday  = todayAtt.filter(function(a){ return a.Status === 'Absent'; }...
[truncated]
[truncated]
[truncated]
-1
+1
[truncated]
[truncated]
-1
+1
[truncated]
[truncated]
