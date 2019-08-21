// Load the Moment.js library once.
var moment = Moment.load()

// Implement String prototype includes
// Cause Google App Scripts is stupid
if (!String.prototype.includes) {
  String.prototype.includes = function (search, start) {
    'use strict'
    if (typeof start !== 'number') {
      start = 0
    }

    if (start + search.length > this.length) {
      return false
    } else {
      return this.indexOf(search, start) !== -1
    }
  }
}

// Get the active spreadsheet
var sheetId = SpreadsheetApp.getActiveSpreadsheet()

// Function to run every hour
function fire () {
  var sheetData = getSheetData()
  if (sheetData.length > 0) {
    createCalendarEvent(sheetData)
  }
}

// Function to obtain data from the sheet
function getSheetData () {
  // Get the last row of the sheet
  var lastRow = sheetId.getLastRow()
  // Get the values of the specified cells
  var data = sheetId.getRange('A2:M' + lastRow).getValues()
  // Get the calendat id from the sheet
  var calendarId = sheetId.getRange('P1').getValues()

  var final = []
  // Iterate through the collected data
  for (var i = 0; i < data.length; i++) {
    if (data[i][8].toLowerCase().includes('back') || data[i][8].toLowerCase().includes('scheduled')) {
      if (!data[i][12].length) {
        final.push({
          calendarId: calendarId[0][0],
          date: data[i][0],
          ext: data[i][1],
          agent: data[i][2],
          url: data[i][3],
          phoneNumber: data[i][4],
          service: data[i][6],
          timeZone: data[i][7],
          statement: data[i][8],
          notes: data[i][9],
          startDate: data[i][10],
          startTime: data[i][11],
          scheduled: data[i][12]
        })
      }
    }
  }
  // return the data
  return final
}

function createCalendarEvent (sheetData) {
  for (var i = 0; i < sheetData.length; i++) {
    var concatStart = moment(sheetData[i].startDate).format('YYYY/MM/DD') + ' ' + moment(sheetData[i].startTime).format('HH:mm')
    var endTime = new Date(moment(concatStart, 'YYYY/MM/DD HH:mm').add(1, 'hour').toDate())

    var description = sheetData[i].notes + '\n\n' +
            'Website URL: ' + sheetData[i].url + '\n' +
            'Phone number: ' + sheetData[i].phoneNumber + '\n' +
            'Service: ' + sheetData[i].service + '\n' +
            'Time zone: ' + sheetData[i].timeZone + '\n' +
            'Call disposition: ' + sheetData[i].statement

    // calendar id stored in the GLOBAL variable object
    var calendar = CalendarApp.getCalendarById(sheetData[i].calendarId)

    // The title for the event that will be created
    var title = 'New appointment ' + sheetData[i].url

    // The start time and date of the event that will be created
    var startTime = new Date(moment(concatStart, 'YYYY/MM/DD HH:mm').toDate())

    // an options object containing the description and guest list
    // for the event that will be created
    var options = {
      description: description,
      guests: 'carlos@topfloormarketing.net',
      sendInvites: true
    }

    try {
      // Set the shceduled flag to true
      sheetId.getRange('M' + (i + 2)).setValue('Y')

      // create a calendar event with given title, start time,
      // end time, and description and guests stored in an
      // options argument
      calendar.getDefaultCalendar().createEvent(title, startTime, endTime, options)
    } catch (e) {
      // create the event without including the guest
      calendar.createEvent(title, startTime, endTime, options)
    }
  }
}
