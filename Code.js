

var caledarIds = ['pwujczyk@google.com']

function executeForToday() {
  execute(0)
}
function executeForYesterday() {
  execute(-1)
}

function executeForLast7Days() {
  for (var e = 7; e >= 0; e--) {
    var day = 0 - e;
    execute(day)
  }
}

function executeForLast30Days() {
  for (var e = 30; e >= 0; e--) {
    var day = 0 - e;
    execute(day)
  }
}

function executeForLast200Days() {
  for (var e = 200; e >= 0; e--) {
    var day = 0 - e;
    execute(day)
  }
}

Date.prototype.getWeekNumber = function () {
  var d = new Date(Date.UTC(this.getFullYear(), this.getMonth(), this.getDate()));
  var dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  var weeknumber = Math.ceil((((d - yearStart) / 86400000) + 1) / 7)
  var yearAndWeek = this.getFullYear() * 100 + weeknumber
  return yearAndWeek
};


function execute(daysOffsetStart) {
  //daysOffsetEnd = 0
  var MINUTE = 60 * 1000;
  var DAY = 24 * 60 * MINUTE;  // ms
  var NOW = new Date();
  NOW.setHours(0, 0, 0, 0);
  var START_DATE = new Date(NOW.getTime() + (daysOffsetStart) * DAY);
  var END_DATE = new Date(NOW.getTime() + (1 + daysOffsetStart) * DAY - MINUTE);


  var start = START_DATE;
  var end = END_DATE;
  clearToday(start, end);
  for (var e = 0; e < caledarIds.length; e++) {
    var calendarId = caledarIds[e];
    processCalendar(calendarId, start, end)
  }
}

function processCalendar(calendarId, start, end) {
  console.log("Hello", start, end)
  var calendar = CalendarApp.getCalendarById(calendarId);
  var calendarName = calendar.getName();
  var events = calendar.getEvents(start, end);
  var day = Utilities.formatDate(start, 'Europe/Warsaw', 'yyyy-MM-dd');
  var weeknumber = start.getWeekNumber()

  var entries = {};
  for (var e = 0; e < events.length; e++) {
    var event = events[e];
    var description = event.getDescription();
    if (description.startsWith('Lunch')) {
      var title = event.getTitle();
      console.log(title);

      var emails = "";
      var people = event.getGuestList();

      for (var i = 0; i < people.length; i++) {
        var email = people[i].getEmail();
        if (email != 'pwujczyk@google.com') {
          emails = email.replace('@google.com', '');
          var dayLog = { day: day, weeknumber: weeknumber, title: title, calendarName: calendarName, emails: emails }
          SaveItem(dayLog)
        }
      }

      var people = event.getCreators();

      for (var i = 0; i < people.length; i++) {

        if (people[i] != 'pwujczyk@google.com') {
          emails = people[i].replace('@google.com', '')
          var dayLog = { day: day, weeknumber: weeknumber, title: title, calendarName: calendarName, emails: emails }
          SaveItem(dayLog)
        }
      }

    }
  }



  return entries;
}


// function columnToLetter(column)
// {
//   var temp, letter = '';
//   while (column > 0)
//   {
//     temp = (column - 1) % 26;
//     letter = String.fromCharCode(temp + 65) + letter;
//     column = (column - temp - 1) / 26;
//   }
//   return letter;
//}

function SaveItem(dayLog) {

  var sheet = getSheet().appendRow([dayLog.day, dayLog.weeknumber, dayLog.title, dayLog.calendarName, dayLog.emails]);
  const lastRow = sheet.getLastRow();
  var xx = sheet.getRange(lastRow, 6);//=SUM(R[-3]C[0]:R[-1]C[0])
  //var letter=columnToLetter(lastRow);
  xx.setValue("=VLOOKUP(E"+lastRow+", 'People map'!$A$2:$B$33, 2, FALSE)")
  //xx.setFormulaR1C1("=VLOOKUP(R["+lastRow+"]C[5], 'People map'!$A$2:$B$33, 2, FALSE')")
//xx.setFormulaR1C1("=VLOOKUP(R[0]C[5], 'People map'!$A$2:$B$33, 2, FALSE')")
   // Logger.log(xx.getValue());

  // const lastColumn = sheet.getLastColumn();
  // const lastCell = sheet.getRange(lastRow, lastColumn);
  // Logger.log(lastCell.getValue());


}

function getSheet() {
  var file = SpreadsheetApp.getActiveSpreadsheet();
  var daily = file.getSheetByName("Lunch");
  return daily;
}

function clearToday(start, end) {
  var sheet = getSheet()
  var data = sheet.getDataRange().getValues();
  for (i = data.length - 1; i > 0; i--) {
    var lineStart = data[i][0]
    var lineEnd = data[i][1]
    if (start <= lineStart && lineStart <= end) {
      console.log(start);
      sheet.deleteRow(i + 1)
    }
  }
}