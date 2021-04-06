/**
 * GLOBAL VARIABLES
 * Client ID and Secret
 */

const clientID = 'cUuYAWnqSzm7Rokou5DeSpHBmS9jh8LcFBatU3QNShuyWpHuJ8NmwG3lAYHSYlIkJ3PMX4DVkxx5gAJO';
const clientSecret = 'yPWF6uemIxKLJLo0h5RMEZJ7B4Vakss1';

function getToken(clientID, clientSecret) {
  //GET TOKEN 
  const oAuthEndpoint = 'https://oauth.oclc.org/token?grant_type=client_credentials&scope=wcapi';

  options = {
    "method": "POST",
    "muteHttpExceptions": true,
    "headers": {
      "Authorization": "Basic " + Utilities.base64Encode(clientID + ':' + clientSecret)
    }
  };

  const response = UrlFetchApp.fetch(oAuthEndpoint, options);
  if (response.getResponseCode() > 200) {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Unable to authenticate: " + response.getContentText())
    Logger.log(response.getResponseCode())
    Logger.log(response.getContentText())
    return null;
  }

  const dataAll = JSON.parse(response.getContentText());
  return dataAll.access_token;
}

function callAPIHoldings(oclcNumber, token) {

  const url = 'https://americas.discovery.api.oclc.org/worldcat/search/v2/bibs-holdings?oclcNumber=' + oclcNumber + '&lat=32.7479&lon=-97.3684&distance=12500&unit=M';

  const options = {
    'method': 'GET',
    'muteHttpExceptions': true,
    'headers': {
      'Authorization': 'Bearer ' + token,
    }
  }

  const responseHoldings = UrlFetchApp.fetch(url, options);
  // Logger.log(responseHoldings.getContentText())
  if (responseHoldings.getResponseCode() > 200) {
    Logger.log("Unable to call the API: " + responseHoldings.getContentText())
    return null;
  }

  return responseHoldings.getContentText();
}


function callAPIBibInfo(oclcNumber, token) {

  const url = 'https://americas.discovery.api.oclc.org/worldcat/search/v2/bibs/' + oclcNumber;

  const options = {
    'method': 'GET',
    'muteHttpExceptions': true,
    'headers': {
      'Authorization': 'Bearer ' + token,
    }
  }

  const responseBibInfo = UrlFetchApp.fetch(url, options);
  // Logger.log(responseBibInfo.getContentText())
  if (responseBibInfo.getResponseCode() > 200) {
    Logger.log("Unable to call the API: " + responseBibInfo.getContentText())
    return null;
  }

  return responseBibInfo.getContentText();
}
