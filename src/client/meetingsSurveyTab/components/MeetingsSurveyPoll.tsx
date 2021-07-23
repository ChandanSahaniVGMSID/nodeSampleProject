
import { Form, FormButton, FormCheckbox, FormInput, Image, Loader } from "@fluentui/react-northstar";
import * as React from "react";
import { IPoll, IQuestion } from "../interfaces";
import GraphService from "../services/GraphService";
import { GetExceptionMessage } from "../utils/MessageUtils";
import { IResponse } from "../interfaces/IResponse";
import { QuestionType } from "../constants/QuestionType";
import { parseBool } from "../utils/Utils";
import strings from "../loc/localizedStrings";
import '../styles/MeetingsSurveyContainer.scss';

const noSurveyImage: any = require('../assets/NoSurveyImage.png');

export interface IMeetingsSurveyPollProps {
    context: microsoftTeams.Context;
    graphService: GraphService;
    poll: IPoll;
    onGetError: (message: string) => void;
}

export interface IMeetingsSurveyPollState {
    isLoading: boolean;
    questions: Array<IQuestion>;
    currentResponses: Array<IResponse>;
    loadedResponses: Array<IResponse>;
}

export class MeetingsSurveyPoll extends React.Component<IMeetingsSurveyPollProps, IMeetingsSurveyPollState> {

    constructor(props) {
        super(props);
        this.state = {
            isLoading: true,
            questions: [],
            currentResponses: [],
            loadedResponses: []
        }

        this.handlers.onLoadPollInformation();
    }

    public render(): React.ReactElement<IMeetingsSurveyPollProps> {

        const { isLoading, questions, loadedResponses } = this.state;


        if (!this.isPollAvailable()) return this._render.emptyPoll();

        const controls = [];
        if (isLoading) {
            controls.push(this._render.loader());
        } else {
            const notAnsweredQuestions = !!questions
                ? questions.filter(question => !loadedResponses.filter(r => r.questionId === question.id)[0])
                : [];

            if (!notAnsweredQuestions || notAnsweredQuestions.length < 1) return this._render.emptyPoll();

            controls.push(
                this._render.content()
            );
        }
        return <div className={"meetingsSurveyPollContainer"}>
            <div className={"meetingsSurveyPollContainerHeader"}>{strings.surveyHeader}</div>
            {controls}
        </div>
    }

    private _render = {
        loader: (): JSX.Element => {
            return <Loader label={strings.loadingPollMessage} />;
        },
        content: (): JSX.Element => {
            const { questions, currentResponses } = this.state;

            return <Form
                onSubmit={this.handlers.onSubmitPoll}
            >
                {questions.map(question => {
                    const responseEntry = currentResponses.filter(r => r.questionId === question.id)[0];
                    switch (question.type) {
                        case QuestionType.YesNo:
                            return <FormCheckbox
                                label={question.title}
                                id={question.id}
                                checked={!!responseEntry ? parseBool(responseEntry.response) : false}
                                onChange={(ev, value) => this.handlers.onQuestionChange(!!value ? value.checked.toString() : "false", question.id)}
                            />;
                        case QuestionType.Text:
                        default:
                            return <FormInput
                                label={question.title}
                                name={question.title}
                                placeholder={strings.questionPlaceholder}
                                id={question.id}
                                required={question.isRequired !== false}
                                value={!!responseEntry ? responseEntry.response : ""}
                                onChange={(ev, value) => this.handlers.onQuestionChange(!!value ? value.value : "", question.id)}
                            />;
                    }
                })}
                <FormButton primary content={strings.submitButtonLabel} />
            </Form>;
        },
        emptyPoll: (): JSX.Element => {
            return <div className="emptyPollContainer">
                <Image src={noSurveyImage?.default} />
                <div className="emptyPollLabel">{strings.emptyPollLabel}</div>
            </div>;
        }
    };

    private handlers = {
        onLoadPollInformation: async () => {
            const { graphService, poll, onGetError } = this.props;

            try {
                if (!this.isPollAvailable()) {
                    return;
                }

                // load questions
                const questionsListPromise = graphService.getQuestionsList(poll.templateId);

                // load responses
                const responsesPromise = graphService.getResponsesForPoll(poll.id);

                const [questionsList, responses] = await Promise.all([questionsListPromise, responsesPromise]);
                await this.setState({
                    isLoading: false,
                    questions: questionsList,
                    currentResponses: responses || [],
                    loadedResponses: responses || []
                });

            } catch (error) {
                this.setState({
                    isLoading: false
                });
                onGetError(GetExceptionMessage(error));
            }
        },
        onQuestionChange: (value, questionId) => {
            const { context, poll } = this.props;
            const { currentResponses } = this.state;

            const newCurrentResponses = currentResponses.slice();
            const responseEntry = newCurrentResponses.filter(r => r.questionId === questionId)[0];
            if (!!responseEntry) {
                responseEntry.response = value;
            } else {
                newCurrentResponses.push({
                    meetingId: context.meetingId,
                    questionId: questionId,
                    response: value,
                    pollId: poll.id,
                    userId: context.userObjectId
                } as IResponse);
            }

            this.setState({
                currentResponses: newCurrentResponses
            });
        },
        onSubmitPoll: async () => {
            const { context, graphService, poll, onGetError } = this.props;
            const { currentResponses, loadedResponses, questions } = this.state;

            try {
                if (!this.isPollAvailable()) {
                    onGetError(strings.pollExpirationMessage);
                }
                await this.setState({
                    isLoading: true,
                });

                const newResponses = currentResponses.filter(cR => !loadedResponses.filter(lR => lR.id === cR.id)[0]);
                const existingResponses = currentResponses.filter(cR => !!loadedResponses.filter(lR => lR.id === cR.id)[0]);

                const nonAnsweredQuestions = questions.filter(q => !currentResponses.filter(r => r.questionId === q.id)[0]);
                newResponses.push(...nonAnsweredQuestions.map(question => {
                    return {
                        meetingId: context.meetingId,
                        questionId: question.id,
                        response: this.getDefaultQuestionResponse(question.type),
                        pollId: poll.id,
                        userId: context.userObjectId
                    } as IResponse;
                }))

                // post responses to a question
                await graphService.postQuestionsResponses(context.meetingId, poll.id, newResponses, existingResponses);

                this.handlers.onLoadPollInformation();

            } catch (error) {
                this.setState({
                    isLoading: false
                });
                onGetError(GetExceptionMessage(error));
            }
        },
    }

    private getDefaultQuestionResponse = (type: QuestionType): string => {
        switch (type) {
            case QuestionType.YesNo:
                return "false";
            case QuestionType.Text:
            default:
                return "";
        }
    }

    private isPollAvailable = (): boolean => {
        const { poll } = this.props;
        if (!poll || !poll.startDateTime || !poll.endDateTime) return false;
        const now = new Date();
        now.setHours(now.getHours() - 1);
        return now < poll.endDateTime;
    }
}