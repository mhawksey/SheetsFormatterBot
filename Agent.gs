function oneOffSetting() { 
  var file = DriveApp.getFilesByName('formatter-a40bd-ad22405d1752.json').next();
  // used by all using this script
  var propertyStore = PropertiesService.getScriptProperties();
  // service account for our Dialogflow agent
  cGoa.GoaApp.setPackage (propertyStore , 
    cGoa.GoaApp.createServiceAccount (DriveApp , {
      packageName: 'dialogflow_serviceaccount',
      fileId: file.getId(),
      scopes : cGoa.GoaApp.scopesGoogleExpand (['cloud-platform']),
      service:'google_service'
    }));
}

/**
 * Detect text message intent from Dialogflow Agent.
 * @param {String} message to find intent
 * @param {String} optLang optional language code
 * @return {object} JSON-formatted response
 */
function handleCommand(message){
  var intent = detectTextMessageIntent(message);
  
  if (!intent.queryResult && !!intent.queryResult.parameters){
    return intent
  }
  var param = intent.queryResult.parameters;
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getActiveSheet();
  var dataRange = sheet.getDataRange();
  
  var values = dataRange.getValues();
  var colors = dataRange.getBackgrounds();
  
  var colIdx = convertCol(param.column);
  var range = param.range || 'cells';
  var color = param.color || 'red';
  var value = param.number;
  var operator = param.operator;
  
  if (range === 'cells'){
    var highlightOp = 'colors[r][colIdx] = color;';
  } else {
    var highlightOp = 'for (var c=0; c < row.length; c++){ colors[r][c] = color; }';
  }
  
  var highlightFn = new Function('colors', 'color', 'row', 'colIdx', 'r', 'if (row[colIdx] '+operator+value+'){'+highlightOp+' }');
  
  // column highlight
  for (var r=0; r<values.length; r++){
    var row = values[r];
    highlightFn(colors, color, row, colIdx, r);
  }  
  dataRange.setBackgrounds(colors);
  return intent;
}

/**
 * Detect text message intent from Dialogflow Agent.
 * @param {String} message to find intent
 * @param {String} optLang optional language code
 * @return {object} JSON-formatted response
 */
function detectTextMessageIntent(message, optLang){
  // setting up calls to Dialogflow with Goa
  var goa = cGoa.GoaApp.createGoa ('dialogflow_serviceaccount',
                                   PropertiesService.getScriptProperties()).execute ();
  if (!goa.hasToken()) {
    throw 'something went wrong with goa - no token for calls';
  }
  // set our token 
  Dialogflow.setTokenService(function(){ return goa.getToken(); } );
   
  /* Preparing the Dialogflow.projects.agent.sessions.detectIntent call 
   * https://cloud.google.com/dialogflow-enterprise/docs/reference/rest/v2/projects.agent.sessions/detectIntent
   *
   * Building a queryInput request object https://cloud.google.com/dialogflow-enterprise/docs/reference/rest/v2/projects.agent.sessions/detectIntent#QueryInput
   * with a TextInput https://cloud.google.com/dialogflow-enterprise/docs/reference/rest/v2/projects.agent.sessions/detectIntent#textinput
  */
  var requestResource = {
    "queryInput": {
      "text": {
        "text": message,
        "languageCode": optLang || "en"
      }
    },
    "queryParams": {
      "timeZone": Session.getScriptTimeZone() // using script timezone but you may want to handle as a user setting
    }
  };
 
 /* Dialogflow.projectsAgentSessionsDetectIntent 
  * @param {string} session Required. The name of the session this query is sent to. Format:`projects//agent/sessions/`.
  * up to the APIcaller to choose an appropriate session ID. It can be a random number orsome type of user identifier (preferably hashed)
  * In this example I'm using for the 
  */
  // your Dialogflow project ID
  var PROJECT_ID = 'formatter-a40bd'; // <- your Dialogflow proejct ID
   
  // using an URI encoded ActiveUserKey (non identifiable) https://developers.google.com/apps-script/reference/base/session#getTemporaryActiveUserKey()
  var SESSION_ID = encodeURIComponent(Session.getTemporaryActiveUserKey()); 
   
  var session = 'projects/'+PROJECT_ID+'/agent/sessions/'+SESSION_ID; // 
  var options = {};
  var intent = Dialogflow.projectsAgentSessionsDetectIntent(session, requestResource, options);
  return intent;
}

function handleAudioCommand(audioUri, optLang){
  // setting up calls to Dialogflow with Goa
  var goa = cGoa.GoaApp.createGoa ('dialogflow_serviceaccount',
                                   PropertiesService.getScriptProperties()).execute ();
  if (!goa.hasToken()) {
    throw 'something went wrong with goa - no token for calls';
  }
  // set our token 
  Dialogflow.setTokenService(function(){ return goa.getToken(); } );
   
  /* Preparing the Dialogflow.projects.agent.sessions.detectIntent call 
   * https://cloud.google.com/dialogflow-enterprise/docs/reference/rest/v2/projects.agent.sessions/detectIntent
   *
   * Building a queryInput request object https://cloud.google.com/dialogflow-enterprise/docs/reference/rest/v2/projects.agent.sessions/detectIntent#QueryInput
   * with a TextInput https://cloud.google.com/dialogflow-enterprise/docs/reference/rest/v2/projects.agent.sessions/detectIntent#textinput
  */
  var requestResource = {
    "queryInput": {
       "audioConfig": {
         "audioEncoding": 'AUDIO_ENCODING_OGG_OPUS',
         "sampleRateHertz": 16000,
         "languageCode": optLang || "en"
      }
    },
    "inputAudio": dataURItoBlob(audioUri),
    "queryParams": {
      "timeZone": Session.getScriptTimeZone() // using script timezone but you may want to handle as a user setting
    }
  };
 Logger.log(dataURItoBlob(audioUri))
 /* Dialogflow.projectsAgentSessionsDetectIntent 
  * @param {string} session Required. The name of the session this query is sent to. Format:`projects//agent/sessions/`.
  * up to the APIcaller to choose an appropriate session ID. It can be a random number orsome type of user identifier (preferably hashed)
  * In this example I'm using for the 
  */
  // your Dialogflow project ID
  var PROJECT_ID = 'formatter-a40bd'; // <- your Dialogflow proejct ID
   
  // using an URI encoded ActiveUserKey (non identifiable) https://developers.google.com/apps-script/reference/base/session#getTemporaryActiveUserKey()
  var SESSION_ID = encodeURIComponent(Session.getTemporaryActiveUserKey()); 
   
  var session = 'projects/'+PROJECT_ID+'/agent/sessions/'+SESSION_ID; // 
  var options = {};
  var intent = Dialogflow.projectsAgentSessionsDetectIntent(session, requestResource, options);
  return intent;
}

function testHighlight() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getActiveSheet();
  var dataRange = sheet.getDataRange();
  
  var values = dataRange.getValues();
  var colors = dataRange.getBackgrounds();
  
  var col = 'D';
  var colIdx = convertCol(col);
  var range = 'rows';
  var color = 'green';
  var value = 1.4;
  var operator = '>';
  if (range === 'cells'){
    var highlightOp = 'colors[r][colIdx] = color;';
  } else {
    var highlightOp = 'for (var c=0; c < row.length; c++){ colors[r][c] = color; }';
  }
  
  var highlightFn = new Function('colors', 'color', 'row', 'colIdx', 'r', 'if (row[colIdx] '+operator+value+'){'+highlightOp+' }');
  
  // column highlight
  for (var r=0; r<values.length; r++){
    var row = values[r];
    highlightFn(colors, color, row, colIdx, r);
  }  
  dataRange.setBackgrounds(colors);
}

// https://stackoverflow.com/a/36949118/1027723
function dataURItoBlob(dataURI, filename) {
  // convert base64/URLEncoded data component to raw binary data held in a string
  var byteString,baseString;
  if (dataURI.split(',')[0].indexOf('base64') >= 0){
    byteString = Utilities.base64Decode(dataURI.split(',')[1]);
    baseString = dataURI.split(',')[1]
  } else {
    byteString = decodeURI(dataURI.split(',')[1]);
  }
  // separate out the mime component
  var mimeString = dataURI.split(',')[0].split(':')[1].split(';')[0];
  return baseString
  return Utilities.newBlob(byteString, mimeString, filename);
}


function reset(){
  // https://stackoverflow.com/a/34350279/1027723
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getActiveSheet();
  var dataRange = sheet.getDataRange();
  dataRange.setBackground(null);
}

function convertCol(val) {
  var base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', i, j, result = 0;

  for (i = 0, j = val.length - 1; i < val.length; i += 1, j -= 1) {
    result += Math.pow(base.length, j) * (base.indexOf(val[i]) + 1);
  }

  return result-1;
};