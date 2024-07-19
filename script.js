function doGet() {
  return HtmlService.createHtmlOutputFromFile('form');
}

function getActivities() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DATA_VALIDATION');
  var activities = sheet.getRange('A2:A').getValues().flat().filter(String);
  return Array.from(new Set(activities));
}

function getSubActivities(activity) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DATA_VALIDATION');
  var data = sheet.getRange('A2:C').getValues();
  var subActivities = data.filter(row => row[0] === activity).map(row => row[1]);
  return Array.from(new Set(subActivities)).filter(String);
}

function getExpenses(subActivity) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DATA_VALIDATION');
  var data = sheet.getRange('A2:C').getValues();
  var expenses = data.filter(row => row[1] === subActivity).map(row => row[2]);
  return Array.from(new Set(expenses)).filter(String);
}

function submitData(formData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var statusSheet = ss.getSheetByName('STATUS');
  
  var lastRow = statusSheet.getLastRow();
  var lastNumber = 0;
  
  if (lastRow > 0) {
    var values = statusSheet.getRange('A1:A' + lastRow).getValues();
    lastNumber = Math.max(...values.flat().filter(value => !isNaN(value)), 0);
  }
  
  var nextNumber = lastNumber + 1;
  statusSheet.appendRow([
    nextNumber, 
    formData.orderNumber,
    formData.activity,
    formData.subActivity,
    formData.expense,
    formData.description,
    formData.invoiceAmount,
    formData.recipient,
    formData.status,
    new Date()
  ]);

  return 'Data submitted successfully!';
}

function getStatusOptions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('STATUS_OPTIONS');
  return sheet.getRange('A2:A').getValues().flat().filter(String);
}

function getSisaAngkas(belanja) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SISA_ANGKAS');
  var data = sheet.getRange('A2:B').getValues();
  
  var match = data.find(row => row[0] === belanja);
  if (match) {
    return formatRupiah(match[1]);
  }
  return 'Tidak ditemukan';
}

function formatRupiah(amount) {
  return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR' }).format(amount);
}
