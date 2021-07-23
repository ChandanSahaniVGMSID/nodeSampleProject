
import { Button, Dialog, Loader, Table } from "@fluentui/react-northstar";
import * as React from "react";
import { IMeeting, IPoll, IQuestion, IUser } from "../interfaces";
import GraphService from "../services/GraphService";
import { GetExceptionMessage } from "../utils/MessageUtils";
import { IResponse } from "../interfaces/IResponse";
import uniq from 'lodash-es/uniq';
import compact from "lodash-es/compact";
import '../styles/MeetingsSurveyContainer.scss';
import strings from "../loc/localizedStrings";
import { QuestionType } from "../constants/QuestionType";

export interface IMeetingsSurveyResultsProps {
    context: microsoftTeams.Context;
    graphService: GraphService;
    meeting: IMeeting;
    poll: IPoll;
    onGetError: (message: string) => void;
}

export interface IMeetingsSurveyResultsState {
    isLoading: boolean;
    questions: Array<IQuestion>;
    currentResponses: Array<IResponse>;
    users: Array<IUser>;
    isLoadingUserInfo: boolean;
}

export class MeetingsSurveyResults extends React.Component<IMeetingsSurveyResultsProps, IMeetingsSurveyResultsState> {

    constructor(props) {
        super(props);
        this.state = {
            isLoading: true,
            questions: [],
            currentResponses: [],
            users: [],
            isLoadingUserInfo: true
        }

        this.handlers.onLoadResultInformation();
    }

    public render(): React.ReactElement<IMeetingsSurveyResultsProps> {
        const { isLoading } = this.state;

        const controls = [];
        if (isLoading) {
            controls.push(this._render.loader());
        } else {
            controls.push(
                this._render.content()
            );
        }
        return <div className={"meetingsSurveyResultsContainer"}>
            <div className={"meetingsSurveyResultsContainerHeader"}>Results</div>
            {controls}</div>
    }

    private _render = {
        loader: (): JSX.Element => {
            return <Loader label={strings.loadingResultsMessage} />;
        },
        content: (): JSX.Element => {
            const { meeting } = this.props;
            const { questions, currentResponses } = this.state;
            if (!questions || questions.length < 1) return <div>{strings.questionsNotFoundMessage}</div>;

            const totalNumberOfAttendees = meeting?.attendees?.length || 0;
            const totalNumberOfParticipants = totalNumberOfAttendees + 1;

            const numberOfParticipantsAnswered = uniq(currentResponses.map(r => r.userId)).length;
            const percentageAnswered = Math.round(numberOfParticipantsAnswered * 100 / totalNumberOfParticipants);
            const responsesString = `${numberOfParticipantsAnswered} people`;
            const percentageString = `${percentageAnswered}% responses`;

            return <div>
                <div className={"meetingsSurveyResultsContainerBody"}>
                    <div>{responsesString}</div>
                    <div>{percentageString}</div>
                </div>
                <div className={"meetingsSurveyResultsContainerActions"}>
                    {this._render.detailsDialog()}
                    {this._render.refreshButton()}
                </div>
            </div>;
        },
        detailsDialog: (): JSX.Element => {
            return <Dialog
                cancelButton={strings.closeButtonLabel}
                className={"meetingsSurveyResultsContainerDialog"}
                content={this._render.detailsContent()}
                header={strings.responsesDetailsHeader}
                trigger={<Button content={strings.viewDetailsButtonLabel} primary />}
            />
        },
        detailsContent: (): JSX.Element => {
            const { questions, currentResponses, isLoadingUserInfo, users } = this.state;
            if (isLoadingUserInfo) return this._render.loader();
            if (!users) return null;
            const usersIds = uniq(currentResponses.map(r => r.userId));


            const header = {
                items: [{ content: <div className="meetingsSurveyResultsContainerTableHeader" title="User name:" >User name:</div> }].concat(...questions.map(question => {
                    return {
                        content: <div className="meetingsSurveyResultsContainerTableHeader" title={question.title} >{question.title}</div>
                    }
                })),
            };

            const rows = compact(usersIds.map((userId, index) => {
                const userInfo = users.filter(user => user.id === userId)[0];
                if (!userInfo || !userInfo.displayName) return null;
                const items = [{ content: <div className="meetingsSurveyResultsContainerTableContent" title={userInfo.displayName}>{userInfo.displayName}</div> }].concat(...questions.map(question => {
                    const response = currentResponses.filter(r => r.userId === userId && r.questionId === question.id)[0];
                    const responseText = !!response && response.response !== null && response.response !== undefined ? response.response : "";

                    return { content: <div className="meetingsSurveyResultsContainerTableContent" title={responseText}>{responseText}</div> };
                }));

                return {
                    key: index,
                    items: items
                };
            }));

            const yesItems = ["Yes"];
            const noItems = ["No"];
            questions.forEach(question => {
                if (question.type !== QuestionType.YesNo) {
                    yesItems.push("-");
                    noItems.push("-");
                } else {
                    const responses = currentResponses.filter(r => r.questionId === question.id);
                    const yesAnswers = responses.filter(r => r.response === "true").length;
                    const noAnswers = responses.length - yesAnswers;
                    yesItems.push(yesAnswers.toString());
                    noItems.push(noAnswers.toString());
                }
            });

            rows.push({
                key: rows.length + 1,
                items: [{ content: <div className="meetingsSurveyResultsContainerTableHeader">Total:</div> }]
            })

            rows.push({
                key: rows.length + 1,
                items: yesItems
            });

            rows.push({
                key: rows.length + 1,
                items: noItems
            });


            return <Table
                className="meetingsSurveyResultsContainerTable"
                header={header}
                rows={rows}
            />;
        },
        refreshButton: (): JSX.Element => {
            return <Button content={strings.resfreshButtonLabel} onClick={this.handlers.onRefreshClick} />;
        }
    };

    private handlers = {
        onLoadResultInformation: async () => {
            const { graphService, poll, onGetError } = this.props;

            try {

                // load questions
                const questionsListPromise = graphService.getQuestionsList(poll.templateId);

                // load responses
                const responsesPromise = graphService.getResponsesForPoll(poll.id, false);

                const [questionsList, responses] = await Promise.all([questionsListPromise, responsesPromise]);

                await this.setState({
                    isLoading: false,
                    questions: questionsList,
                    currentResponses: responses || []
                });

                this.handlers.onLoadUserInfo();

            } catch (error) {
                this.setState({
                    isLoading: false,
                    isLoadingUserInfo: false
                });
                onGetError(GetExceptionMessage(error));
            }
        },
        onLoadUserInfo: async () => {
            const { graphService, onGetError } = this.props;
            const { currentResponses } = this.state;

            try {
                const usersIds = uniq(currentResponses.map(r => r.userId));
                const users = await graphService.getUsersInfo(usersIds);


                await this.setState({
                    users: users || [],
                    isLoadingUserInfo: false
                });

            } catch (error) {
                this.setState({
                    isLoadingUserInfo: false
                });
                onGetError(GetExceptionMessage(error));
            }
        },
        onRefreshClick: async () => {
            await this.setState({
                isLoading: true
            })
            this.handlers.onLoadResultInformation();
        }
    }
}