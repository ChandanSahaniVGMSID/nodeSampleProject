import * as React from "react";
import { Provider, Flex, Text, Loader } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import AuthenticationService from "./services/AuthenticationService";
import { MeetingsSurveyContainer } from "./components/MeetingsSurveyContainer";
import { GetExceptionMessage } from "./utils/MessageUtils";
import strings from "./loc/localizedStrings";
import { LocalStorageService } from "./services/LocalStorageService";
import { ConfigurationKeys } from "./constants";
import ConfigurationService from "./services/ConfigurationService";
import './styles/MeetingsSurveyContainer.scss';
import { isInMeetingPanel } from "./utils/Utils";

/**
 * Implementation of the Meetings Survey content page
 */
export const MeetingsSurveyTab = () => {

    const [{ inTeams, theme, context }] = useTeams();
    const [config, setConfig] = useState<object>({});
    const [aadToken, setAadToken] = useState<string>("");
    const [error, setError] = useState<string>();
    const [isLoading, setIsLoading] = useState<boolean>(true);
    const [isInMeeting, setIsInMeeting] = useState<boolean>(false);

    const handleError = (message: string) => {
        setError(message);
        setIsLoading(false);
        microsoftTeams.appInitialization.notifyFailure({
            reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
            message
        });
    }

    useEffect(() => {
        if (context) {

            setIsInMeeting(isInMeetingPanel(context));

            // load configuration (environment variables) from node server
            LocalStorageService.Config.getConfig(() => ConfigurationService.getConfig(), true)
                .then(config => {
                    setConfig(config);

                    microsoftTeams.authentication.getAuthToken({
                        successCallback: async (idToken: string) => {
                            try {
                                const authService = new AuthenticationService(context, config);
                                const aadToken = await LocalStorageService.AuthToken.getAuthToken(context.userObjectId, () => authService.getServerToken(idToken), true);

                                if (!aadToken) {
                                    handleError(strings.aadTokenNotValidMessage);
                                } else {
                                    setAadToken(aadToken);
                                    setIsLoading(false);
                                    microsoftTeams.appInitialization.notifySuccess();
                                }
                            } catch (error) {
                                handleError(GetExceptionMessage(error));
                            }
                        },
                        failureCallback: (message: string) => {
                            handleError(message);
                        },
                        resources: [config[ConfigurationKeys.AppRegistration.AppUri]]
                    });
                })
                .catch(error => handleError(GetExceptionMessage(error)));
        }

    }, [context]);

    const classNames = ["appContainer"];
    if (isInMeeting) {
        classNames.push("inMeetingContainer");
    }

    return (
        <Provider
            className={classNames.join(" ")}
            theme={theme}
        >
            {inTeams
                ? <Flex fill={true} column styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}>
                    <Flex.Item>
                        {
                            isLoading
                                ? <Loader label={strings.loadingAuthenticationTokenMessage} />
                                : !aadToken
                                    ? <div><Text content={`${strings.ssoErrorMessage} ${error}`} /></div>
                                    : <MeetingsSurveyContainer
                                        aadToken={aadToken}
                                        context={context}
                                        config={config}
                                    />
                        }
                    </Flex.Item>
                </Flex>
                : <div>{strings.teamsContextOnlyMessage}</div>}
        </Provider>
    );
};
