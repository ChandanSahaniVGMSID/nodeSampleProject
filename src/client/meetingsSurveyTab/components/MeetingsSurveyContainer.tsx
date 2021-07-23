
import { Alert, Loader } from "@fluentui/react-northstar";
import * as React from "react";
import { MeetingRole } from "../constants/MeetingRole";
import { IMeeting } from "../interfaces/IMeeting";
import { IMessage } from "../interfaces/IMessage";
import { IPoll } from "../interfaces/IPoll";
import GraphService from "../services/GraphService";
import { ConcatMessage, GetExceptionMessage } from "../utils/MessageUtils";
import { MeetingsSurveyTemplatesList } from "./MeetingsSurveyTemplatesList";
import '../styles/MeetingsSurveyContainer.scss';
import { MeetingsSurveyPoll } from "./MeetingsSurveyPoll";
import { MeetingsSurveyResults } from "./MeetingsSurveyResults";
import strings from "../loc/localizedStrings";
import { isInMeetingPanel } from "../utils/Utils";
import { AppConstants } from "../constants";

export interface IMeetingsSurveyContainerProps {
    aadToken: string;
    context: microsoftTeams.Context;
    config: object;
}

export interface IMeetingsSurveyContainerState {
    isLoading: boolean;
    loadingMessage: string;
    messages: Array<IMessage>;
    meeting: IMeeting;
    poll: IPoll;
}

export class MeetingsSurveyContainer extends React.Component<IMeetingsSurveyContainerProps, IMeetingsSurveyContainerState> {

    private graphService: GraphService;

    constructor(props) {
        super(props);
        this.state = {
            isLoading: true,
            loadingMessage: strings.loadingMeetngInfoMessage,
            messages: [],
            meeting: null,
            poll: null
        }
    }

    public componentDidMount() {
        const { aadToken, context, config } = this.props;
        try {
            this.graphService = new GraphService(aadToken, context, config);
        } catch (error) {
            this.setState({
                isLoading: false,
                loadingMessage: "",
                messages: ConcatMessage(this.state.messages, {
                    text: GetExceptionMessage(error)
                }),
            });
        }
        this.handlers.onLoadMeetingInformation();
    }

    public render(): React.ReactElement<IMeetingsSurveyContainerProps> {
        const { isLoading } = this.state;
        const controls = [];
        if (isLoading) {
            controls.push(this._render.loader());
        } else {
            controls.push(
                this._render.messages(),
                this._render.content()
            );
        }

        return <div className="meetingsSurveyContainer">
            {controls}
        </div>;
    }

    private _render = {
        loader: (): JSX.Element => {
            const { loadingMessage } = this.state;
            return <Loader label={loadingMessage} />;
        },
        messages: (): JSX.Element => {
            const { messages } = this.state;

            if (!messages || messages.length < 1) return null;

            return <div>{
                messages.map(m => <Alert
                    danger
                    dismissible
                    dismissAction={{
                        'aria-label': 'close',
                    }}
                    content={m.text}
                />)
            }</div>;
        },
        content: (): JSX.Element => {
            const role: MeetingRole = this.getCurrentUserMeetingRole();
            let template: JSX.Element = null;
            switch (role) {
                case MeetingRole.Organizer:
                    template = this._render.organizerUserContent();
                    break;
                case MeetingRole.Attendee:
                default:
                    template = this._render.attendeeUserContent();
                    break;
            }
            return <div>{template}</div>;
        },
        organizerUserContent: (): JSX.Element => {
            const { context } = this.props;
            const { poll, meeting } = this.state;
            const isInMeeting = isInMeetingPanel(context);

            // if the poll is not created for that meeting yet display template selection screen (in pre-meeting mode only)
            if (!poll) {
                return isInMeeting
                    ? <MeetingsSurveyPoll
                        context={context}
                        graphService={this.graphService}
                        poll={poll}
                        onGetError={this.handlers.onGetError}
                    />
                    : <MeetingsSurveyTemplatesList
                        context={context}
                        graphService={this.graphService}
                        onTemplateLaunched={this.handlers.onTemplateLaunched}
                        onGetError={this.handlers.onGetError}
                    />;
            }

            // display result window and poll window for organizer
            return <div>
                {!isInMeeting && <MeetingsSurveyResults
                    context={context}
                    graphService={this.graphService}
                    meeting={meeting}
                    poll={poll}
                    onGetError={this.handlers.onGetError}
                />}
                <MeetingsSurveyPoll
                    context={context}
                    graphService={this.graphService}
                    poll={poll}
                    onGetError={this.handlers.onGetError}
                />
            </div>;
        },
        attendeeUserContent: (): JSX.Element => {
            const { context } = this.props;
            const { poll } = this.state;

            // display poll window for attendee
            return <MeetingsSurveyPoll
                context={context}
                graphService={this.graphService}
                poll={poll}
                onGetError={this.handlers.onGetError}
            />;
        },
    };

    private handlers = {
        onLoadMeetingInformation: async () => {
            const { context } = this.props;

            try {

                // load meeting details
                const meetingDetailsPromise = this.graphService.getMeetingDetails();

                // load meeting poll details
                const pollDetailsPromise = this.graphService.getMeetingPoll(context.meetingId);

                const [meetingDetails, pollDetails] = await Promise.all([meetingDetailsPromise, pollDetailsPromise]);
                this.setState({
                    isLoading: false,
                    loadingMessage: "",
                    meeting: meetingDetails,
                    poll: pollDetails
                });

            } catch (error) {
                this.setState({
                    isLoading: false,
                    loadingMessage: "",
                    messages: ConcatMessage(this.state.messages, {
                        text: GetExceptionMessage(error)
                    }),
                });
            }
        },
        onTemplateLaunched: async (templateId: string) => {
            const { context } = this.props;
            const { meeting } = this.state;

            try {
                await this.setState({
                    isLoading: true,
                    loadingMessage: strings.settingTemplateMessage,
                });

                // create meeting poll
                const meetingPollPromise = this.graphService.createMeetingPoll(templateId, context.meetingId, meeting);

                // get tabId for link
                const tabIdPromise = this.graphService.getMeetingTabId(AppConstants.TabName);

                const [meetingPoll, tabId] = await Promise.all([meetingPollPromise, tabIdPromise]);

                if (!meetingPoll) {
                    this.setState({
                        isLoading: false,
                        loadingMessage: "",
                        messages: ConcatMessage(this.state.messages, {
                            text: strings.pollNotCreatedMessage
                        }),
                    });
                    return;
                }

                this.graphService.postChatMessage(tabId);

                this.setState({
                    isLoading: false,
                    loadingMessage: "",
                    poll: meetingPoll
                });

            } catch (error) {
                this.setState({
                    isLoading: false,
                    loadingMessage: "",
                    messages: ConcatMessage(this.state.messages, {
                        text: GetExceptionMessage(error)
                    }),
                });
            }
        },
        onGetError: (message: string) => {
            this.setState({
                messages: ConcatMessage(this.state.messages, {
                    text: message
                }),
            });
        }
    }

    private getCurrentUserMeetingRole = (): MeetingRole => {
        const { context } = this.props;
        const { meeting } = this.state;

        if (meeting?.organizer?.id === context.userObjectId) return MeetingRole.Organizer;

        return MeetingRole.Attendee;
    }
}
