var fs = require('fs');
var google = require('googleapis');
var googleAuth = require('google-auth-library');

function clockIn(token, timestamp, docId) {
  return new Promise(function (resolve, reject) {
    authorize(token).then(function (client) {
      sheetsIn(client, docId, timestamp).then(function () {
        resolve({docId: docId});
      });
    }).catch(function (err) {
      reject(err);
    });
  });
}

function clockOut(token, timestamp, docId) {
  return new Promise(function (resolve, reject) {
    authorize(token).then(function (client) {
      sheetsOut(client, docId, timestamp).then(function () {
        resolve({docId: docId});
      });
    }).catch(function (err) {
      reject(err);
    });
  });
}

function createSpreadsheet(token, docId) {
  return new Promise(function (resolve, reject) {
    authorize(token).then(function (client) {
      createDocument(client, docId).then(function (docId) {
        resolve(docId);
      });
    }).catch(function (err) {
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
  return new Promise(function (resolve, reject) {
    // Load client secrets from a local file.
    fs.readFile('client_secret.json', function processClientSecrets(err, content) {
      if (err) {
        reject('Error loading client secret file: ' + err)
      }
      // Authorize a client with the loaded credentials, then call the
      // Google Sheets API.
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
 * Print the names and majors of students in a sample spreadsheet:
 * https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
 */
function sheetsIn(auth, docId, timestamp) {
  return new Promise(function (resolve, reject) {
    var date = new Date(parseInt(timestamp));
    var sheets = google.sheets('v4');
    sheets.spreadsheets.values.append({
      auth: auth,
      spreadsheetId: docId,
      range: 'Sheet1!A1:B1',
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [[date.toLocaleDateString(), date.toLocaleTimeString()]]
      }
    }, function(err, response) {
      if (err) {
        console.log('The API returned an error: ' + err);
        return reject('The API returned an error: ' + err);
      }
      resolve('Finished!');
    });
  });
}

/**
 * Print the names and majors of students in a sample spreadsheet:
 * https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
 */
function sheetsOut(auth, docId, timestamp) {
  return new Promise(function (resolve, reject) {
    var date = new Date(parseInt(timestamp));
    var sheets = google.sheets('v4');
    sheets.spreadsheets.values.append({
      auth: auth,
      spreadsheetId: docId,
      range: 'Sheet1!A1:B1',
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [[date.toLocaleDateString(), date.toLocaleTimeString()]]
      }
    }, function(err, response) {
      if (err) {
        console.log('The API returned an error: ' + err);
        return reject('The API returned an error: ' + err);
      }
      resolve('Finished!');
    });
  });
}

function createDocument(auth, docId) {
  return new Promise(function (resolve, reject) {
    if (docId === '-1') {
      var sheets = google.sheets('v4');
      sheets.spreadsheets.create({
        auth: auth,
        resource: {
          properties: {
            title: 'CLOCKY TIMESHEET'
          }
        },
        sheets: [
          {
            resource: {
              properties: {
                title: 'TIMESHEET'
              }
            }
          }, {
            resource: {
              properties: {
                title: 'PAYMENT'
              }
            }
          }
        ]
      }, function (err, response) {
        if (err) {
          reject({code: err.code, message: err.message});
        } else {
          resolve(response.spreadsheetId);
        }
      }).catch(function (err) {
        console.log(err);
      });
    } else {
      resolve(docId);
    }
  });
}

module.exports = {
  clockIn: clockIn,
  clockOut: clockOut,
  createSpreadsheet: createSpreadsheet
};