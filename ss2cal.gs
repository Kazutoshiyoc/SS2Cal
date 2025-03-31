1// =====================================================================================
//
//                               SpreadSheet to Calendar
//
// =====================================================================================
const spreadsheet_id = 'spreadsheetidspreadsheetidspreadsheetid';
const sheet_name     = 'シート1';
const calendar_id = 'example@gmail.com';

// ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function saveEnv(ssVals,calVals) {
  var env_ss = {
    ID:             ssVals[0],
    SHEET_NAME:     ssVals[1],
    SHEET_LAST_ROW: ssVals[2]
  };
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET',JSON.stringify(env_ss));

  var env_cal = {
    ID: calVals[0]
  };
  PropertiesService.getScriptProperties().setProperty('CALENDAR',JSON.stringify(env_cal));
}
function getEnv() {
  var env_ss_raw = PropertiesService.getScriptProperties().getProperty('SPREADSHEET');
  var env_ss = JSON.parse(env_ss_raw);
  var env_cal_raw = PropertiesService.getScriptProperties().getProperty('CALENDAR');
  var env_cal = JSON.parse(env_cal_raw);
  return [env_ss, env_cal];
}
// ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Spread Sheet
function getSpreadSheetByID(id) {
  var ss = SpreadsheetApp.openById(id);
  return ss;
}
function getSheetByName(ss,name) {
  var sheet = ss.getSheetByName(name);
  return sheet;
}
function getEventByRange(sheet,range) {
  var area = sheet.getRange(range);
  var events = area.getValues();
  return events;
}
// ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Calendar
function createEventSchedule(id, event) {
  var description = `ID: ${event[0]}, ${event[5]}`;
  var calendar = CalendarApp.getCalendarById(id);
  calendar.createEvent(       // createEvent(title, startTime, endTime, options)
    event[3],                   // Event Title
    new Date(event[1]),         // Event Start Date
    new Date(event[2]),         // Event End Date
    {                           // Options
      location: event[4],       // Location
      description: description  // Description
    }
  );
}
// ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function init() {
  saveEnv(
    [spreadsheet_id, sheet_name, '1'],
    [calendar_id]
  );
}
function main() {
  var [SPREADSHEET, CALENDAR] = getEnv();                                       //Logger.log(SPREADSHEET.ID);
                                                                                //Logger.log(SPREADSHEET.SHEET_NAME);
                                                                                //Logger.log(SPREADSHEET.SHEET_LAST_ROW);
                                                                                //Logger.log(CALENDAR.ID);
  var spreadsheet = getSpreadSheetByID(SPREADSHEET.ID);                         //Logger.log(spreadsheet.getName());
  var sheet = getSheetByName(spreadsheet,SPREADSHEET.SHEET_NAME);               //Logger.log(sheet.getName());
  var currentLastRow = sheet.getLastRow();                                      //Logger.log(currentLastRow);
  var prevLastRow = SPREADSHEET.SHEET_LAST_ROW;                                 //Logger.log(prevLastRow);
  if (currentLastRow > prevLastRow) {
    var events = getEventByRange(sheet,`A${currentLastRow}:F${prevLastRow}`);   //Logger.log(events);
    for (var i = prevLastRow; i < events.length; i++) {
      createEventSchedule(CALENDAR.ID, events[i]);                              Logger.log(events[i]);
    }
  }
  saveEnv(
    [spreadsheet_id, sheet_name, currentLastRow],
    [calendar_id]
  );
}