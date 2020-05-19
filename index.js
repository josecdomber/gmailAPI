const fs = require('fs');
const readline = require('readline');
const { google } = require('googleapis');
const json2xls = require('json2xls');

let totalMessages = 0;

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/gmail.readonly'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';

console.log('Comienza la descarga ...');

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
  if (err) return console.log('Error loading client secret file:', err);
  // Authorize a client with credentials, then call the Gmail API.
  //authorize(JSON.parse(content), listLabels);
  //authorize(JSON.parse(content), getLabel);
  authorize(JSON.parse(content), listMessages);
});




/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  const { client_secret, client_id, redirect_uris } = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
    client_id, client_secret, redirect_uris[0]);

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback);
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client);
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
      if (err) return console.error('Error retrieving access token', err);
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) return console.error(err);
        console.log('Token stored to', TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  });
}


/**
 * Lists the labels in the user's account.
 *
 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
 */
/*  Número de mensajes de una etiqueta */
function getLabel(auth) {
  const gmail = google.gmail({ version: 'v1', auth });
  gmail.users.labels.get({
    userId: 'me',
    id: 'Label_6460148555393810731'
  }, (err, res) => {
    if (err) return console.log('The API returned an error: ' + err);
    const label = res.data;
    console.log('Objeto label: ', label);
    if (label) {
      console.log(`- ${label.name} - Mensajes en total: ${label.messagesTotal}`);
    } else {
      console.log('No labels found.');
    }
  });
}


/*  Listado de todas las etiquetas */
function listLabels(auth) {
  const gmail = google.gmail({ version: 'v1', auth });
  gmail.users.labels.list({
    userId: 'me',
  }, (err, res) => {
    if (err) return console.log('The API returned an error: ' + err);
    const labels = res.data.labels;
    if (labels.length) {
      console.log('Labels:');
      labels.forEach((label) => {
        console.log(`- ${label.name} - Id: ${label.id}`);
      });
    } else {
      console.log('No labels found.');
    }
  });
}

async function listMessages(auth) {
  const gmail = google.gmail({ version: 'v1', auth });
  var getPageOfMessages = async (result, totalResult) => {
    // console.log('Resultado parcial: ', result.data);

    totalResult = totalResult.concat(result.data.messages);
    var nextPageToken = result.data.nextPageToken;
    if (nextPageToken) {
      result = await gmail.users.messages.list({
        'userId': 'me',
        'pageToken': nextPageToken,
        'labelIds': ['Label_6460148555393810731']
      });
      getPageOfMessages(result, totalResult);
    } else {
      // console.log('List  messages: ', totalResult);
      // console.log('Total messages: ', totalResult.length);
      showEachMessage(auth, totalResult);
    }

  };
  var initialResult = await gmail.users.messages.list({
    'userId': 'me',
    'labelIds': ['Label_6460148555393810731']
  });
  getPageOfMessages(initialResult, []);
}

async function showEachMessage(auth, totalMessages) {
  //console.log('List  messages showEachMessage: ', totalResult);
  //console.log('Total messages showEachMessage: ', totalResult.length);
  const gmail = google.gmail({ version: 'v1', auth });
  let arrayMensajes = [];
  let objMensaje = {};

  for (let i = 0; i < totalMessages.length; i++) {
    let message = totalMessages[i];
    let request = await gmail.users.messages.get({
      'userId': 'me',
      'id': message.id,
      'format': 'full'
    });

    objMensaje.From = request.data.payload.headers.find(item => item.name === 'From').value;
    objMensaje.To = request.data.payload.headers.find(item => item.name === 'To').value;
    objMensaje.Date = request.data.payload.headers.find(item => item.name === 'Date').value;
    objMensaje.Subject = request.data.payload.headers.find(item => item.name === 'Subject').value;
    objMensaje.Body = request.data.snippet;

    objMensaje.Nombre   = extraerStringEntreDosString(request.data.snippet, 'Nombre:', 'Email:');
    objMensaje.Email    = extraerStringEntreDosString(request.data.snippet, 'Email:', 'Teléfono:');
    if (objMensaje.Email === null) {
      objMensaje.Email = extraerStringEntreDosString(request.data.snippet, 'Email:', 'Solicitud');
    }
    objMensaje.Telefono = extraerStringEntreDosString(request.data.snippet, 'Teléfono:', 'Solicitud');


    arrayMensajes.push(objMensaje);

    objMensaje = {};

    console.log('Contador de mensajes: ', i + 1);
  }
  
  // Escribo el fichero JSON
  // fs.appendFile('Mensajes.json', arrayMensajes, function (err) {
  //   if (err) throw err;
  // });

  var xls = json2xls(arrayMensajes);

  fs.writeFileSync('Mensajes.xlsx', xls, 'binary');

}

function extraerStringEntreDosString(cadena, strA, strB) {
  const posIniStrA = cadena.search(strA);
  const posFinStrA = posIniStrA === -1 ? 0 : posIniStrA + strA.length;

  if (posFinStrA === 0) return null;

  const posIniStrB = cadena.search(strB);

  if (posIniStrB === -1) return null;

  return cadena.slice(posFinStrA, posIniStrB);

}


//
// Fuera de uso
async function writeFile(request) {

  fs.appendFile('Mensajes.txt', `
    ____________________________________________________________________
    ___ INICIO _________________________________________________________ 
    From ${request.data.payload.headers.find(item => item.name === 'From').value}
    To ${request.data.payload.headers.find(item => item.name === 'To').value}
    Date ${request.data.payload.headers.find(item => item.name === 'Date').value}
    Subject ${request.data.payload.headers.find(item => item.name === 'Subject').value}
    ${request.data.snippet}
    ___   FIN  _________________________________________________________
    ____________________________________________________________________`
    , function (err) {
      if (err) throw err;
    });

}

