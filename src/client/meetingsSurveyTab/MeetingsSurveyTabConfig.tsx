import * as React from "react";
import { Provider, Flex, Header } from "@fluentui/react-northstar";
import { useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import { GetGuid } from "./utils/Utils";
import { AppConstants } from "./constants";
import strings from "./loc/localizedStrings";

/**
 * Implementation of Meetings Survey configuration page
 */
export const MeetingsSurveyTabConfig = () => {

    const [{ theme, context }] = useTeams({});

    const onSaveHandler = (saveEvent: microsoftTeams.settings.SaveEvent) => {
        const host = "https://" + window.location.host;

        microsoftTeams.settings.setSettings({
            contentUrl: host + "/meetingsSurveyTab/?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}",
            websiteUrl: host + "/meetingsSurveyTab/?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}",
            suggestedDisplayName: AppConstants.TabName,
            removeUrl: host + "/meetingsSurveyTab/remove.html?theme={theme}",
            entityId: GetGuid()
        });
        saveEvent.notifySuccess();
    };

    useEffect(() => {
        if (context) {
            microsoftTeams.settings.registerOnSaveHandler(onSaveHandler);
            microsoftTeams.settings.setValidityState(true);
            microsoftTeams.appInitialization.notifySuccess();
        }
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [context]);

    return (
        <Provider theme={theme}>
            <Flex fill={true}>
                <Flex.Item>
                    <div>
                        <Header content={strings.configurationMessage} />
                    </div>
                </Flex.Item>
            </Flex>
        </Provider>
    );
};
