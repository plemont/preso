var CLIENT_ID = 'INSERT_CLIENT_ID';
var CLIENT_SECRET = 'INSERT_CLIENT_SECRET';

/**
 * This script allows the stepping through of the Authorization Code Grant in
 * order to obtain a refresh token.
 *
 * This script uses the out-of-band redirect URI, which is not part of the
 * OAuth2 standard, to allow not redirecting the user. If this does not work
 * with your API, try instead the OAuth playground:
 * https://developers.google.com/oauthplayground/
 *
 * Execute script twice:
 * Execution 1, will result in a URL, which when placed in the browser will
 * issue a code.
 * Execution 2: place the code in "CODE" below and execute. If successful a
 * refresh token will be printed to the console.
 */

// Enter required scopes, e.g. ['https://www.googleapis.com/auth/drive']
var SCOPES = ['https://www.googleapis.com/auth/presentations',
              'https://www.googleapis.com/auth/drive'];

// Auth URL, e.g. https://accounts.google.com/o/oauth2/auth
var AUTH_URL = 'https://accounts.google.com/o/oauth2/auth';
// Token URL, e.g. https://accounts.google.com/o/oauth2/token
var TOKEN_URL = 'https://accounts.google.com/o/oauth2/token';

var CODE = '';

function main() {
  if (CODE) {
    generateRefreshToken();
  } else {
    generateAuthUrl();
  }
}

/**
 * Creates the URL for pasting in the browser, which will generate the code
 * to be placed in the CODE variable.
 */
function generateAuthUrl() {
  var payload = {
    scope: SCOPES.join(' '),
    // Specify that no redirection should take place
    // This is Google-specific and not part of the OAuth2 specification.
    redirect_uri: 'urn:ietf:wg:oauth:2.0:oob',
    response_type: 'code',
    access_type: 'offline',
    client_id: CLIENT_ID
  };
  var options = {payload: payload};
  var request = UrlFetchApp.getRequest(AUTH_URL, options);
  Logger.log(
      'Browse to the following URL: ' + AUTH_URL + '?' + request.payload);
}

/**
 * Generates a refresh token given the authorization code.
 */
function generateRefreshToken() {
  var payload = {
    code: CODE,
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    // Specify that no redirection should take place
    // This is Google-specific and not part of the OAuth2 specification.
    redirect_uri: 'urn:ietf:wg:oauth:2.0:oob',
    grant_type: 'authorization_code'
  };
  var options = {method: 'POST', payload: payload};
  var response = UrlFetchApp.fetch(TOKEN_URL, options);
  var data = JSON.parse(response.getContentText());
  if (data.refresh_token) {
    var msg = 'Success! Refresh token: ' + data.refresh_token +
      '\n\nThe following may also be a useful format for pasting into your script:\n\n' +
      'var CLIENT_ID = \'' + CLIENT_ID + '\';\n' +
      'var CLIENT_SECRET = \'' + CLIENT_SECRET + '\';\n' +
      'var REFRESH_TOKEN = \'' + data.refresh_token + '\';\n';
    Logger.log(msg);
  } else {
    Logger.log(
        'Error, failed to generate Refresh token: ' +
        response.getContentText());
  }
}
