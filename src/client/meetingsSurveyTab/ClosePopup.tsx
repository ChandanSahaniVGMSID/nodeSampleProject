import * as React from "react";
import { useEffect } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import strings from "./loc/localizedStrings";

/**
 * This component is loaded when the Azure implicit grant flow has completed.
 */
export const ClosePopup = () => {
    
    //Helper function that converts window.location.hash into a dictionary
    const getHashParameters = () => {
        let hashParams: any = {};
        window.location.hash.substr(1).split("&").forEach(function (item) {
            let [key, value] = item.split('=');
            hashParams[key] = decodeURIComponent(value);
        });
        return hashParams;
    }

    useEffect(() => {
        microsoftTeams.initialize();

        //The Azure implicit grant flow injects the result into the window.location.hash object. Parse it to find the results.
        let hashParams = getHashParameters();

        //If consent has been successfully granted, the Graph access token should be present as a field in the dictionary.
        if (hashParams["access_token"]) {
            //Notifify the showConsentDialogue function in Tab.js that authorization succeeded. The success callback should fire. 
            microsoftTeams.authentication.notifySuccess(hashParams["access_token"]);
        } else {
            microsoftTeams.authentication.notifyFailure(strings.consentFailedMessage);
        }
    });

    return (
        <div>
            <h1>{strings.consentFlowCompletedMessage}</h1>
        </div>
    );
}

export default ClosePopup;