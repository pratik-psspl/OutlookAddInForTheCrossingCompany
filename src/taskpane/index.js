import "office-ui-fabric-react/dist/css/fabric.min.css";
import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import * as React from "react";
import * as ReactDOM from "react-dom";
import * as OfficeHelpers from "@microsoft/office-js-helpers";
import { StaticConst } from "./helper/Const";
import { AsyncHelper } from "./helper/AsyncHelper";
initializeIcons();

let isOfficeInitialized = false;
let accessToken;

const title = "The Crossing Group";
var helper = new AsyncHelper();
const render = (Component) => {
  ReactDOM.render(
    <AppContainer>
      <Component title={title} isOfficeInitialized={isOfficeInitialized} Authorization={accessToken} />
    </AppContainer>,
    document.getElementById("container")
  );
};
var clark = (function () {
  let saveAttachmentsRequest;
  var init=function(){
    startAuthenticationProcess();
  }
  var officeHelperAuthenticationGetAccessToken = function () {
    var authenticator = new OfficeHelpers.Authenticator();
    if (OfficeHelpers.Authenticator.isAuthDialog()) return;
    authenticator.endpoints.registerMicrosoftAuth(StaticConst.ActiveDirectory.Id, {
     // responseType: "code", //code
      redirectUrl: StaticConst.ActiveDirectory.redirectUrl,
      scope: StaticConst.ActiveDirectory.scope,
    });
    authenticator = new OfficeHelpers.Authenticator();
    authenticator
      .authenticate(OfficeHelpers.DefaultEndpoints.Microsoft, false)
      .then(function (token) {        
        renderAppComponentWithToken(token.access_token);      
      })
      .catch(function (e) {
        // callback(false);
        OfficeHelpers.Utilities.log(e, "Error on authentication");
      });
  };
  var officeHelperAuthenticationGetCode = function () {
    var authenticator = new OfficeHelpers.Authenticator();
    if (OfficeHelpers.Authenticator.isAuthDialog()) return;
    authenticator.endpoints.registerMicrosoftAuth(StaticConst.ActiveDirectory.Id, {
      responseType: "code", //code
      redirectUrl: StaticConst.ActiveDirectory.redirectUrl,
      scope: StaticConst.ActiveDirectory.scope,
    });
    authenticator = new OfficeHelpers.Authenticator();
    authenticator
      .authenticate(OfficeHelpers.DefaultEndpoints.Microsoft, false)
      .then(function (token) {
        const saveAttachmentsRequest = {
          auth_code: token.code,
          refresh_token: null,
        };
        Office.context.roamingSettings.set(StaticConst.roamingStorageName, saveAttachmentsRequest);
        Office.context.roamingSettings.saveAsync();

        startAuthenticationProcess();
      })
      .catch(function (e) {
        // callback(false);
        OfficeHelpers.Utilities.log(e, "Error on authentication");
      });
  };
  var renderAppComponentWithToken = function (token) {
    accessToken = `Bearer ${token}`;
    isOfficeInitialized = true;
    render(App);
  };
  var startAuthenticationProcess = function () {   
    const roamingStoredData = Office.context.roamingSettings.get(StaticConst.roamingStorageName);
    saveAttachmentsRequest = {
      auth_code: roamingStoredData ? roamingStoredData.auth_code : null,
      refresh_token: roamingStoredData ? roamingStoredData.refresh_token : null,
      RedirectUri: StaticConst.ActiveDirectory.redirectUrl,
      AppId: StaticConst.ActiveDirectory.Id,
      AppPassword: StaticConst.ActiveDirectory.ClientSecret,
    };

    if (saveAttachmentsRequest.auth_code != null || saveAttachmentsRequest.refresh_token != null) {
      helper
        .postToAzureFunction(JSON.stringify(saveAttachmentsRequest))
        .then((token) => {
          if (token.data.access_token) {
            const saveAttachmentsRequest = {
              auth_code: token.data.access_token,
              refresh_token: token.data.refresh_token,
            };
            Office.context.roamingSettings.set(StaticConst.roamingStorageName, saveAttachmentsRequest);
            Office.context.roamingSettings.saveAsync();
            renderAppComponentWithToken(token.data.access_token);
          } else {
            officeHelperAuthenticationGetCode();
          }
        })
        .catch((ex) => {         
          officeHelperAuthenticationGetAccessToken();          
        });
    } else {
      if (StaticConst.AccessToken) {
        accessToken = `Bearer ${StaticConst.AccessToken}`;
        isOfficeInitialized = true;
        render(App);
      } else {
        officeHelperAuthenticationGetCode();
      }
    }
  };
  return {
    init:init,renderAppComponentWithToken:renderAppComponentWithToken
  };
})();

/* Render application after Office initializes */
Office.initialize = () => {
  
  if(StaticConst.AccessToken){
    clark.renderAppComponentWithToken(StaticConst.AccessToken);
  }
  else{
  clark.init();}
};

/* Initial render showing a progress bar */
render(App);
if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
