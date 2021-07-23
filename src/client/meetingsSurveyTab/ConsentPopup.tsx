// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import * as React from "react";
import { GetGuid, getGraphScope } from "./utils/Utils";
import { AppConstants } from "./constants";
import strings from "./loc/localizedStrings";

/**
 * This component is used to redirect the user to the Azure authorization endpoint from a popup
 */
export const ConsentPopup = () => {

  const queryString = window.location.search;
  const urlParams = new URLSearchParams(queryString);
  const tenantId = urlParams.get(AppConstants.UrlParameters.TenantId);
  const AppId = urlParams.get(AppConstants.UrlParameters.AppId);

  //Form a query for the Azure implicit grant authorization flow
  //https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-implicit-grant-flow      
  let queryParams: any = {
    tenant: `${tenantId}`,
    client_id: `${AppId}`, //Client ID of the Azure AD app registration ( may be from different tenant for multitenant apps)
    response_type: "token", //token_id in other samples is only needed if using open ID
    scope: getGraphScope(),
    redirect_uri: window.location.origin + "/meetingsSurveyTab/auth-end",
    nonce: GetGuid()
  }

  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?`;
  queryParams = new URLSearchParams(queryParams).toString();
  const authorizeEndpoint = url + queryParams;

  //Redirect to the Azure authorization endpoint. When that flow completes, the user will be directed to auth-end
  //Go to ClosePopup.js
  window.location.assign(authorizeEndpoint);

  return (
    <div>
      <h1>{strings.redirectToConsentPageMessage}</h1>
    </div>
  );
}

export default ConsentPopup;
