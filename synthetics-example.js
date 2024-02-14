// See documentation here - https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/overview?view=dataverse-latest
// Postman example - https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-postman-perform-operations?view=dataverse-latest
// Querying data https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query-data-web-api

var assert = require('assert');

// Scope of the OAuth access request
var scope = "https://company.api.crm11.dynamics.com/.default";
// OAuth request URI, uses the Azure Directory (tenant) id secret (AZURE_DIRECTORY_ID).
var oauthUIR = "https://login.microsoftonline.com/" + $secure.AZURE_DIRECTORY_ID + "/oauth2/v2.0/token";
// Dynamics API endpoint
var dynamicsQueryURI = "https://company.api.crm11.dynamics.com/api/data/v9.2/accounts?$select=name,accountnumber&$top=3";
// New Relic event API endpoint. Uses a account id secret (NR_ACCOUNT_ID).
// Please ensure you use the correct endpoint EU or US (insights-collector.newrelic.com or insights-collector.eu01.nr-data.net).
var nRPostEndPoint = "https://insights-collector.eu01.nr-data.net/v1/accounts/" + $secure.NR_ACCOUNT_ID + "/events";
// New Relic Event name
var nREventType = "CRMAccount";

// Function to post results back into New Relic.
// Uses a secret called NR_INSERT_LICENSE_KEY for the New Relic insert key.
function postNRData(jsonBody) {
  var nROptions = {
    uri: nRPostEndPoint,
    headers: {
      'Api-Key': $secure.INSIGHTS_KEY,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(jsonBody)
  };
  
  $http.post(nROptions,
    function (error, response, body) {
      assert.equal(response.statusCode, 200, 'Expected a 200 OK response from New Relic');
      console.log('Event data sent to New Relic.');
    }
  );
}

// Function to get the first 3 accounts from MS Dynamics 365.
function getAccounts(token, url) {
  var dynamicsOptions = {
    'uri': url,
    'headers': {
    'Authorization': 'Bearer ' + token
    }
  };

  $http.get(dynamicsOptions,
      function (err, response, body) {
        // JSON object to send to New Relic, empty for now!
        var jsonNR = [{}];
        // Did we get the right response code back from Azure?
        assert.equal(response.statusCode, 200, 'Expected a 200 OK response');
        // Parse the response into JSON
        var jsonBody = JSON.parse(body);
        // Walk through each of the entity details returned
        for (var i = 0; i < jsonBody.value.length; i++) {
          // Build a JSON object for this entity and start filling it in
          var js = {};
          js.eventType = nREventType;

          var jsPos = jsonBody.value[i];
          js.accountName = jsPos.name;
          js.accountNumber = jsPos.accountnumber;
          js.accountId = jsPos.accountid;

          // Add the JSON object we have built for this entity to the overall object we will send to New Relic.
          jsonNR[i] = js;
          console.log('accountName: ' + js.accountName + ', accountNumber: ' + js.accountNumber + ', accountId: ' + js.accountId);
        }
        console.log(jsonBody.value.length + ' records returned by Dynamics.');
        // Post the data to New Relic.
        postNRData(jsonNR);
        // If there is another page of data, then get it.
        // Note - this has not been tested!!
        if (jsonBody.hasOwnProperty("@odata") && jsonBody["@odata"].hasOwnProperty("nextLink")) {
          // Call this function recursively.
          getAccounts(token, String(jsonBody["@odata"].nextLink));
        }
      }
  );
};

// OAuth configuration options, sends form data to Microsoft login.
// Uses CRM_CLIENT_ID and CRM_CLIENT_SECRET secrets for the client id and secret.
var OAuthOptions = {
  'uri': oauthUIR,
  form: {
    'scope': scope,
    'client_id': $secure.CRM_CLIENTID,
    'client_secret':$secure.CRM_CLIENT_SECRET,
    'grant_type': 'client_credentials'
  },
  headers: {
    'Content-Type': 'application/x-www-form-urlencoded',
    'Accept': 'application/json'
  }
};

// Get the OAuth token
$http.post(OAuthOptions,
  function (err, response, body) {
    assert.equal(response.statusCode, 200, 'Expected a 200 OK response');
    // Parse the JSON and get the token.
    var jsRtn = JSON.parse(body);
    assert.equal(jsRtn.hasOwnProperty('access_token'), true, 'Access token not found.');
    console.log('OAuth token found.');
    // Now get the health data.
    getAccounts(jsRtn.access_token, dynamicsQueryURI);
  }
);
