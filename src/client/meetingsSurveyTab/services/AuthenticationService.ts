import * as microsoftTeams from "@microsoft/teams-js";
import { AppConstants, ConfigurationKeys } from "../constants";

export default class AuthenticationService {

    constructor(private context: microsoftTeams.Context, private config: object) {
        if (!context) throw new Error("Invalid context");
        if (!config) throw new Error("Invalid config");
    }

    public getServerToken = async (token: string): Promise<string> => {

        const serverURL = `${window.location.origin}/getGraphAccessToken?ssoToken=${token}`;
        const response = await fetch(serverURL).catch(this.unhandledFetchError); //This calls getGraphAccessToken route in /api-server/app.js
        if (response) {
            const data = await response.json().catch(this.unhandledFetchError);

            if (!response.ok && data.error === 'consent_required') {
                //A consent_required error means it's the first time a user is logging into to the app, so they must consent to sharing their Graph data with the app.
                //They may also see this error if MFA is required.
                this.showConsentDialog();


            } else if (!response.ok) {
                //Unknown error                
            } else {
                //Server side token exchange worked. Save the access_token to state, so that it can be picked up and used by the componentDidMount lifecycle method.
                return data['access_token'];
            }
        }
        return "";
    }

    private showConsentDialog = () => {

        microsoftTeams.authentication.authenticate({
            url: window.location.origin + `/meetingssurveytab/auth-start.html?${AppConstants.UrlParameters.TenantId}=${this.context.tid}&${AppConstants.UrlParameters.AppId}=${this.config[ConfigurationKeys.AppRegistration.AppId]}`,
            width: 600,
            height: 535,
            successCallback: (result) => {
                console.log("showConsentDialog success" + result);
            },
            failureCallback: (reason) => {
                console.log("showConsentDialog error" + reason);
            }
        });
    }

    private unhandledFetchError = (err: string) => {
        console.error("Unhandled fetch error: ", err);
    }
}