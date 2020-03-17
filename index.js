'use strict'

const fs = require('fs');
const readline = require('readline');
const {google} = require('googleapis');
const googleSheetID = process.env.GOOGLE_SHEET_ID;

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = __dirname + '/token.json';

exports.handler = (event, context, callback) => {
  console.log("Event: \n" + JSON.stringify(event, null, 2));

  // Load client secrets from a local file.
  return fs.readFile('credentials.json', (err, content) => {
    if (err) return console.log('Error loading client secret file:', err);

    // Authorize a client with credentials, then call the Google Sheets API.
    authorize(JSON.parse(content), updateCellValues, event.target, callback);
  });
};

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback, target, maincallback) {
  const {client_secret, client_id, redirect_uris} = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
    client_id, client_secret, redirect_uris[0]);

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback);
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client, target, maincallback);
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  console.log('Authorize this app by visiting this url:', authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.question('Enter the code from that page here: ', (code) => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return console.error('Error while trying to retrieve access token', err);
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) return err;
        return 'Token stored to:'+ TOKEN_PATH;
      });
      callback(oAuth2Client);
    });
  });
}

/**
 * Prints the values in bannerman spreadsheet:
 * @see https://docs.google.com/spreadsheets/d/1zKjNeebjO55aKwbnb6Pjtsofu4nRSulPXUkG3EMEhus/edit
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */
function updateCellValues(auth, target, mailcallback) {
  const sheets = google.sheets({version: 'v4', auth});
  sheets.spreadsheets.values.get({
    spreadsheetId: googleSheetID,
    range: `Values!${target}`,
  }, (err, res) => {
    if (err) return console.log('The API returned an error: ' + err);

    const rows = res.data.values;
    if (rows.length) {
      let values = [[Number(rows[0])+1]];
      const resource = {
        values: values
      };
      sheets.spreadsheets.values.update({
        spreadsheetId: googleSheetID,
        range: `Values!${target}`,
        valueInputOption: 'RAW',
        resource: resource,
      }, (err, res) => {
        if (err) return console.log('The API returned an error: ' + err);

        if (res.status == 200) {
          console.log('Values updated');
          mailcallback(null, res.status);
        } else {
          console.log('Update failed.');
          mailcallback(null, res.status);
        }
      });
    } else {
      console.log('No data found.');
      mailcallback(null, res.status);
    }
  });
}
