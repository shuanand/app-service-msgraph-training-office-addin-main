// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <ConsentJsSnippet>
'use strict';

const msalClient = new msal.PublicClientApplication({
  auth: {
    clientId: authConfig.clientId
  }
});

const msalRequest = {
  scopes: [
    'https://graph.microsoft.com/.default'
  ]
};

// Function that handles the redirect back to this page
// once the user has signed in and granted consent
function handleResponse(response) {
  localStorage.removeItem('msalCallbackExpected');
  if (response !== null) {
    localStorage.setItem('msalAccountId', response.account.homeId);
    Office.context.ui.messageParent(JSON.stringify({ status: 'success', result: response.accessToken }));
  }
}

Office.initialize = function () {
  if (Office.context.ui.messageParent) {
    // Let MSAL process a redirect response if that's what
    // caused this page to load.
    msalClient.handleRedirectPromise()
      .then(handleResponse)
      .catch((error) => {
        console.log(error);
        Office.context.ui.messageParent(JSON.stringify({ status: 'failure', result: error }));
      });

    // If we're not expecting a callback (because this is
    // the first time the page has loaded), then start the
    // login process
    if (!localStorage.getItem('msalCallbackExpected')) {
      // Set the msalCallbackExpected property so we don't
      // make repeated token requests
      localStorage.setItem('msalCallbackExpected', 'yes');

      // If the user has signed into this machine before
      // do a token request, otherwise do a login
      if (localStorage.getItem('msalAccountId')) {
        msalClient.acquireTokenRedirect(msalRequest);
      } else {
        msalClient.loginRedirect(msalRequest);
      }
    }
  }
}
// </ConsentJsSnippet>
