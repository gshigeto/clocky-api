var fs = require('fs');
var google = require('googleapis');
var googleAuth = require('google-auth-library');
var dateFormat = require('dateformat');

/**
 * Wrapper to authorize and clock in on Clocky Timesheet
 *
 * @param {Object} token Object containing the access_token, refresh_token
 * and token_type used to create OAuth2 client credentials
 * @param {date} timestamp Timestamp used to clock the user out
 * @param {string} docId Google Spreadsheet ID
*/
function clockIn(token, timestamp, docId) {
  return new Promise((resolve, reject) => {
    authorize(token).then(client => {
      sheetsIn(client, docId, timestamp).then(response => {
        resolve(response);
      }).catch(err => {
        reject(err);
      });
    }).catch(err => {
      reject(err);
    });
  });
}

/**
 * Wrapper to authorize and clock out on Clocky Timesheet
 *
 * @param {Object} token Object containing the access_token, refresh_token
 * and token_type used to create OAuth2 client credentials
 * @param {date} timestamp Timestamp used to clock the user out
 * @param {string} docId Google Spreadsheet ID
*/
function clockOut(token, timestamp, docId) {
  return new Promise((resolve, reject) => {
    authorize(token).then(client => {
      sheetsOut(client, docId, timestamp).then(response => {
        resolve(response);
      }).catch(err => {
        reject(err);
      });
    }).catch(err => {
      reject(err);
    });
  });
}

/**
 * Wrapper to authorize and export shifts
 *
 * @param {Object} token Object containing the access_token, refresh_token
 * and token_type used to create OAuth2 client credentials
 * @param {Object[]} shifts An array of shifts to export to Google Sheets
 * @param {string} docId Google Spreadsheet ID
*/
function exportTimesheet(token, shifts, docId) {
  return new Promise((resolve, reject) => {
    authorize(token).then(client => {
      exportToSheets(client, docId, shifts).then(response => {
        resolve(response);
      }).catch(err => {
        reject(err);
      });
    }).catch(err => {
      reject(err);
    });
  });
}

/**
 * Wrapper to authorize and create a Google Spreadsheet
 *
 * @param {Object} token Object containing the access_token, refresh_token
 * and token_type used to create OAuth2 client credentials
 * @param {string} docId Google Spreadsheet ID
*/
function createSpreadsheet(token, docId) {
  return new Promise((resolve, reject) => {
    authorize(token).then(client => {
      createDocument(client, docId).then(docId => {
        resolve({code: 200, body: {message: 'Spreadsheet successfully created', docId: docId}});
      }).catch(err => {
        reject(err);
      });
    }).catch(err => {
      reject(err);
    });
  });
}

/**
 * Authorize with the given token
 *
 * @param {Object} token The authorization client tokens.
 */
function authorize(token) {
  return new Promise((resolve, reject) => {
    // Load client secrets from a local file.
    fs.readFile('client_secret.json', function processClientSecrets(err, content) {
      if (err) {
        reject({code: 400, message: 'Could not load client_secret.json'});
      }
      // Authorize a client with the loaded credentials
      resolve(createOauthClient(JSON.parse(content), token));
    });
  });
}

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 *
 * @param {Object} credentials The authorization client credentials.
 * @param {Object} token The authorization client tokens.
 */
function createOauthClient(credentials, token, callback, docId) {
  var clientSecret = credentials.installed.client_secret;
  var clientId = credentials.installed.client_id;
  var redirectUrl = credentials.installed.redirect_uris[0];

  var auth = new googleAuth();
  var oauth2Client = new auth.OAuth2(clientId, clientSecret, redirectUrl);
  oauth2Client.credentials = token;
  return oauth2Client;
}

/**
 * Clocks the user in and sets the current date
 *
 * @param {Object} auth OAuth2 client credentials
 * @param {string} docId Google Spreadsheet ID
 * @param {date} timestamp Timestamp used to clock the
 * user in as well as set the date
*/
function sheetsIn(auth, docId, timestamp) {
  return new Promise((resolve, reject) => {
    var date = new Date(timestamp);
    var sheets = google.sheets('v4');
    sheets.spreadsheets.values.append({
      auth: auth,
      spreadsheetId: docId,
      range: 'Timesheet!A1:B1',
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [[date.toDateString(), dateFormat(date, 'h:MM:ss TT')]]
      }
    }, (err, response) => {
      if (err) {
        reject({code: err.code, message: err.message});
      } else {
        resolve({code: 200, body: {message: 'Successfully clocked in', docId: docId}});
      }
    });
  });
}

/**
 * Clocks the user out and sets the total time worked for that shift
 *
 * @param {Object} auth OAuth2 client credentials
 * @param {string} docId Google Spreadsheet ID
 * @param {date} timestamp Timestamp used to clock the user out
*/
function sheetsOut(auth, docId, timestamp) {
  return new Promise((resolve, reject) => {
    getCurrentRow(auth, docId).then(row => {
      var date = new Date(timestamp);
      var sheets = google.sheets('v4');
      sheets.spreadsheets.values.update({
        auth: auth,
        spreadsheetId: docId,
        range: `Timesheet!C${row}:E${row}`,
        valueInputOption: 'USER_ENTERED',
        resource: {
          values: [[dateFormat(date, 'h:MM:ss TT'), '', `=IF(B${row}<>"",IF(C${row}="","MISSING OUT",C${row}-B${row}),IF(C${row}<>"", "MISSING IN", ""))`]]
        }
      }, (err, response) => {
        if (err) {
          reject({code: err.code, message: err.message});
        } else {
          resolve({code: 200, body: {message: 'Successfully clocked in', docId: docId}});
        }
      });
    }).catch(err => {
      reject(err);
    });
  });
}

/**
 * Clocks the user out and sets the total time worked for that shift
 *
 * @param {Object} auth OAuth2 client credentials
 * @param {string} docId Google Spreadsheet ID
 * @param {Object[]} shifts An array of shifts to export to Google Sheets
*/
function exportToSheets(auth, docId, shifts) {
  return new Promise((resolve, reject) => {
    clearSheet(auth, docId).then(() => {
      createShiftValues(shifts).then(values => {
        var sheets = google.sheets('v4');
        sheets.spreadsheets.values.update({
          auth: auth,
          spreadsheetId: docId,
          range: `Timesheet!A2:E`,
          valueInputOption: 'USER_ENTERED',
          resource: {
            values: values
          }
        }, (err, response) => {
          if (err) {
            reject({code: err.code, message: err.message});
          } else {
            resolve({code: 200, body: {message: 'Successfully exported to Sheets', docId: docId}});
          }
        });
      }).catch(err => {
        reject(err);
      });
    }).catch(err => {
      reject(err);
    });
  });
}

/**
 * Clears current values from A2:E
 *
 * @param {Object} auth OAuth2 client credentials
 * @param {string} docId Google Spreadsheet ID
*/
function clearSheet(auth, docId) {
  return new Promise((resolve, reject) => {
    var sheets = google.sheets('v4');
    sheets.spreadsheets.values.clear({
      auth: auth,
      spreadsheetId: docId,
      range: 'Timesheet!A2:E'
    }, (err, response) => {
      if (err) {
        reject({code: err.code, message: err.message});
      } else {
        resolve();
      }
    });
  });
}

/**
 * Creates an array of arrays for each row
 *
 * @param {Object[]} shifts An array of shifts to export to Google Sheets
*/
function createShiftValues(shifts) {
  return new Promise((resolve, reject) => {
    var values = [];
    for (var i = 0; i < shifts.length; i++) {
      var clockIn = new Date(shifts[i].clockIn);
      var clockOut = shifts[i].clockOut != undefined ? dateFormat(new Date(shifts[i].clockOut), 'h:MM:ss TT') : '';
      values.push([clockIn.toDateString(), dateFormat(clockIn, 'h:MM:ss TT'), clockOut, '', `=IF(B${i+2}<>"",IF(C${i+2}="","MISSING OUT",C${i+2}-B${i+2}),IF(C${i+2}<>"", "MISSING IN", ""))`]);
    }
    resolve(values);
  });
}

/**
 * Gets the current row that is empty from range C1:E
 *
 * @param {Object} auth OAuth2 client credentials
 * @param {string} docId Google Spreadsheet ID
*/
function getCurrentRow(auth, docId) {
  return new Promise((resolve, reject) => {
    var sheets = google.sheets('v4');
    sheets.spreadsheets.values.get({
      auth: auth,
      spreadsheetId: docId,
      range: 'Timesheet!C1:E'
    }, (err, response) => {
      if (err) {
        reject({code: err.code, message: err.message});
      } else {
        if (response.values) {
          resolve(response.values.length + 1);
        } else {
          resolve(1);
        }
      }
    });
  });
}

/**
 * Creates a new Google Spreadsheet named 'Clocky Timesheet'
 * with a sheet titled 'Timesheet'
 *
 * @param {Object} auth OAuth2 client credentials
 * @param {string} docId Google Spreadsheet ID
*/
function createDocument(auth, docId) {
  return new Promise((resolve, reject) => {
    if (docId === '-1') {
      var sheets = google.sheets('v4');
      sheets.spreadsheets.create({
        auth: auth,
        resource: {
          properties: {
            title: 'Clocky Timesheet'
          },
          sheets: [
            {
              properties: {
                title: 'Timesheet'
              }
            }
          ]
        },
      }, (err, response) => {
        if (err) {
          reject({code: err.code, message: err.message});
        } else {
          createHeader(auth, response.spreadsheetId).then(() => {
            createTotalsFormulas(auth, response.spreadsheetId).then(() => {
              resolve(response.spreadsheetId);
            }).catch(err => {
              reject(err);
            });
          }).catch(err => {
              reject(err);
          });
        }
      });
    } else {
      resolve(docId);
    }
  });
}

/**
 * Creates header for the sheet including: date, in, out, total, and note
 *
 * @param {Object} auth OAuth2 client credentials
 * @param {string} docId Google Spreadsheet ID
*/
function createHeader(auth, docId) {
  return new Promise((resolve, reject) => {
    var sheets = google.sheets('v4');
    sheets.spreadsheets.values.append({
      auth: auth,
      spreadsheetId: docId,
      range: 'Timesheet!A1:G1',
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [
          ['Date', 'In', 'Out', '', 'Total', '', 'NOTE: You must currently change column E to duration number format as well as the \'total\' cell'],
        ]
      }
    }, (err, response) => {
      if (err) {
        reject({code: err.code, message: err.message});
      } else {
        resolve();
      }
    });
  });
}

/**
 * Creates the start and end date cells along with the formula
 * to calculate the total time worked between the two dates
 *
 * @param {Object} auth OAuth2 client credentials
 * @param {string} docId Google Spreadsheet ID
*/
function createTotalsFormulas(auth, docId) {
  return new Promise((resolve, reject) => {
    var date = new Date().toISOString();
    var sheets = google.sheets('v4');
    sheets.spreadsheets.values.append({
      auth: auth,
      spreadsheetId: docId,
      range: 'Timesheet!G2:H4',
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [
          ['Start Date:', date.toDateString()],
          ['End Date:', date.toDateString()],
          ['Total Time:', '=SUMIFS(E2:E,A2:A,">="&H2,A2:A,"<="&H3)']
        ]
      }
    }, (err, response) => {
      if (err) {
        reject({code: err.code, message: err.message});
      } else {
        resolve();
      }
    });
  });
}

// Currently not correct
function setDurationType(auth, docId) {
  return new Promise((resolve, reject) => {
    var date = new Date().toISOString();
    var sheets = google.sheets('v4');
    sheets.spreadsheets.values.update({
      auth: auth,
      spreadsheetId: docId,
      range: 'Timesheet!E2:E',
      resource: {
          "type": "DATE",
          "pattern": "[hh]:[mm]:[ss]"
      }
    }, (err, response) => {
      if (err) {
        reject({code: err.code, message: err.message});
      } else {
        resolve('Finished!');
      }
    });
  });
}

module.exports = {
  clockIn: clockIn,
  clockOut: clockOut,
  exportTimesheet: exportTimesheet,
  createSpreadsheet: createSpreadsheet
};