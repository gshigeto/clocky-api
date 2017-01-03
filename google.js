var fs = require('fs');
var google = require('googleapis');
var googleAuth = require('google-auth-library');
var Q = require('q');

function clockIn(token, checkId) {
  var promise = Q.defer();
  authorize(token).then(function (client) {
    createDocument(client, checkId).then(function (docId) {
      listMajors(client, docId).then(function () {
        promise.resolve({docId: docId});
      });
    });
  }).catch(function (err) {
    promise.reject(err);
  });
  return promise.promise;
}

function clockOut(token, checkId) {
  var promise = Q.defer();
  authorize(token).then(function (client) {
    createDocument(client, checkId).then(function (docId) {
      listMajors(client, docId).then(function () {
        if (docId !== checkId) {
          promise.resolve({docId: docId});
        }
      });
    });
  }).catch(function (err) {
    promise.reject(err);
  });
  return promise.promise;
}

/**
 * Authorize with the given token
 *
 * @param {Object} token The authorization client tokens.
 */
function authorize(token) {
  var promise = Q.defer();
  // Load client secrets from a local file.
  fs.readFile('client_secret.json', function processClientSecrets(err, content) {
    if (err) {
      promise.reject('Error loading client secret file: ' + err)
    }
    // Authorize a client with the loaded credentials, then call the
    // Google Sheets API.
    promise.resolve(createOauthClient(JSON.parse(content), token));
  });
  return promise.promise;
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
function listMajors(auth, docId) {
  var promise = Q.defer();
  var sheets = google.sheets('v4');
  sheets.spreadsheets.values.get({
    auth: auth,
    spreadsheetId: '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms',
    range: 'Class Data!A2:E',
  }, function(err, response) {
    if (err) {
      console.log('The API returned an error: ' + err);
      promise.reject('The API returned an error: ' + err);
    }
    var rows = response.values;
    if (rows.length == 0) {
      console.log('No data found.');
      promise.resolve('No data found.');
    } else {
      console.log('Name, Major:');
      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        // Print columns A and E, which correspond to indices 0 and 4.
        console.log('%s, %s', row[0], row[4]);
      }
      promise.resolve('Finished!');
    }
  });
  return promise.promise;
}

function createDocument(auth, docId) {
  var promise = Q.defer();
  if (docId === '-1') {
    var sheets = google.sheets('v4');
    sheets.spreadsheets.create({
      auth: auth
    }, function (err, response) {
      if (err) {
        console.log('The API returned an error: ' + err);
        promise.reject('The API returned an error: ' + err);
      }
      promise.resolve(response.spreadsheetId);
    })
  } else {
    promise.resolve(docId);
  }
  return promise.promise;
}

module.exports = {
  clockIn: clockIn,
  clockOut: clockOut
};