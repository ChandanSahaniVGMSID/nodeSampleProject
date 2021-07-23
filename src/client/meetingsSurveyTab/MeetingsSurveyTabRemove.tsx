import * as React from "react";
import { Provider, Flex, Text, Header } from "@fluentui/react-northstar";
import { useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import strings from "./loc/localizedStrings";

/**
 * Implementation of Meetings Survey remove page
 */
export const MeetingsSurveyTabRemove = () => {

    const [{ inTeams, theme }] = useTeams();

    useEffect(() => {
        if (inTeams === true) {
            microsoftTeams.appInitialization.notifySuccess();
        }
    }, [inTeams]);

    return (
        <Provider theme={theme}>
            <Flex fill={true}>
                <Flex.Item>
                    <div>
                        <Header content={strings.removeTabMessage} />
                        <Text content={strings.removeTabAdditionalMessage} />
                    </div>
                </Flex.Item>
            </Flex>
        </Provider>
    );
};
