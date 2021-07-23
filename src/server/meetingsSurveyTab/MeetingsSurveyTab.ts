import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/meetingsSurveyTab/index.html")
@PreventIframe("/meetingsSurveyTab/config.html")
@PreventIframe("/meetingsSurveyTab/remove.html")
export class MeetingsSurveyTab {
}
