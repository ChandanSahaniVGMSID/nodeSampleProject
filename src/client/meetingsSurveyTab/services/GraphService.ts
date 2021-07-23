import compact from "lodash-es/compact";
import { ConfigurationKeys } from "../constants";
import { QuestionType } from "../constants/QuestionType";
import { IQuestion, IUser } from "../interfaces";
import { IMeeting, IMeetingParticipant } from "../interfaces/IMeeting";
import { IPoll } from "../interfaces/IPoll";
import { IResponse } from "../interfaces/IResponse";
import { ITemplate } from "../interfaces/ITemplate";
import { GetGuid } from "../utils/Utils";

export default class GraphService {
    constructor(private aadToken: string, private context: microsoftTeams.Context, private config: object) {
        if (!aadToken) throw new Error("Invalid aadToken");
        if (!context) throw new Error("Invalid context");
        if (!config) throw new Error("Invalid config");

        this.initUrls();
    }

    private siteDomain: string = "";
    private serverRelativeUrl: string = "";

    public getMeetingDetails = async (): Promise<IMeeting> => {
        const url = `https://graph.microsoft.com/v1.0/me/onlineMeetings?$filter=JoinWebUrl%20eq%20'https://teams.microsoft.com/l/meetup-join/${this.context?.chatId}/0?context={"Tid":"${this.context?.tid}","Oid":"${this.context?.userObjectId}"}'`;

        let result;
        try {
            result = await this.getData(url);
        } catch (error) {
            // if the user is not organizer he will get error while trying to get meeting
            // hence we log an error (at this time is was something like 'The query specified in the URI is not valid. The requested resource is not a collection. Query options $filter, $orderby, $count, $skip, and $top can be applied only on collections.')
            // for now we need meeting object only to indentify the role of a user. so if meeting is empty we'll consider user as attendde
            // if later we will need meeting info for something else we could consider caching it to SP list
            console.log(error);
        }

        if (!result || !result.value || !result.value[0]) {
            return null;
        };

        const { value } = result;
        const item = value[0];
        const { participants } = item;
        const organizer: IMeetingParticipant = {
            upn: participants?.organizer?.upn,
            id: participants?.organizer?.identity?.user?.id
        };
        const attendees: Array<IMeetingParticipant> = participants?.attendees?.map(item => {
            return {
                upn: item?.upn,
                id: item?.identity?.user?.id
            }
        });
        const meeting = {
            id: item["id"],
            startDateTime: new Date(item["startDateTime"]),
            endDateTime: new Date(item["endDateTime"]),
            subject: item["subject"],
            organizer: organizer,
            attendees: attendees
        } as IMeeting;

        return meeting;
    }

    public getTemplatesList = async (): Promise<Array<ITemplate>> => {

        const titleFieldName = this.config[ConfigurationKeys.Lists.Templates.Fields.Title];
        const descriptionFieldName = this.config[ConfigurationKeys.Lists.Templates.Fields.Description];
        const questionsFieldName = this.config[ConfigurationKeys.Lists.Templates.Fields.Questions];

        const listTitle = this.config[ConfigurationKeys.Lists.Templates.Title];
        const selectedFields: Array<string> = compact([titleFieldName, descriptionFieldName, questionsFieldName]);
        const url = `https://graph.microsoft.com/v1.0/sites/${this.siteDomain}:${this.serverRelativeUrl}:/lists/${listTitle}/items?expand=fields(select=${selectedFields.join(',')})`;

        const result = await this.getData(url);
        if (!result || !result.value) return [];
        const templates: Array<ITemplate> = compact(result.value.map(item => {
            if (!item || !item.fields) return null;
            return {
                id: item["id"],
                title: item.fields[titleFieldName],
                description: item.fields[descriptionFieldName],
                questionsIds: !!item.fields[questionsFieldName]
                    ? item.fields[questionsFieldName].map(i => i["LookupId"])
                    : []
            } as ITemplate;
        }));
        return templates;
    }

    public getMeetingPoll = async (meetingId: string): Promise<IPoll> => {

        const meetingIdFieldName = this.config[ConfigurationKeys.Lists.Polls.Fields.MeetingId];
        const meetingOrganizerFieldName = this.config[ConfigurationKeys.Lists.Polls.Fields.MeetingOrganizer];
        const meetingAttendeesFieldName = this.config[ConfigurationKeys.Lists.Polls.Fields.MeetingAttendees];
        const templateFieldName = this.config[ConfigurationKeys.Lists.Polls.Fields.Template];
        const startDateFieldName = this.config[ConfigurationKeys.Lists.Polls.Fields.StartDate];
        const endDateFieldName = this.config[ConfigurationKeys.Lists.Polls.Fields.EndDate];

        const listTitle = this.config[ConfigurationKeys.Lists.Polls.Title];
        const selectedFields: Array<string> = compact([meetingIdFieldName, meetingOrganizerFieldName, meetingAttendeesFieldName, templateFieldName, startDateFieldName, endDateFieldName]);
        const url = `https://graph.microsoft.com/v1.0/sites/${this.siteDomain}:${this.serverRelativeUrl}:/lists/${listTitle}/items?expand=fields(select=${selectedFields.join(',')})&$filter=fields/${meetingIdFieldName} eq '${meetingId}'`;

        const result = await this.getData(url, true);
        if (!result || !result.value || !result.value[0]) return null;
        const { value } = result;
        const item = value[0];

        const poll = {
            id: item["id"],
            meetingId: item.fields[meetingIdFieldName],
            meetingOrganizer: item.fields[meetingOrganizerFieldName],
            meetingAttendees: item.fields[meetingAttendeesFieldName],
            templateId: item.fields[templateFieldName],
            startDateTime: new Date(item.fields[startDateFieldName]),
            endDateTime: new Date(item.fields[endDateFieldName]),
        } as IPoll;

        return poll;
    }

    public createMeetingPoll = async (templateId: string, meetingId: string, meeting: IMeeting): Promise<IPoll> => {
        if (!templateId) throw new Error("GraphService: createMeetingPoll - Empty templateId");
        if (!meetingId) throw new Error("GraphService: createMeetingPoll - Empty meetingId");
        if (!meeting) throw new Error("GraphService: createMeetingPoll - Empty meeting object");

        const listTitle = this.config[ConfigurationKeys.Lists.Polls.Title];
        const url = `https://graph.microsoft.com/v1.0/sites/${this.siteDomain}:${this.serverRelativeUrl}:/lists/${listTitle}/items`;

        const meetingIdFieldName = this.config[ConfigurationKeys.Lists.Polls.Fields.MeetingId];
        const meetingOrganizerFieldName = this.config[ConfigurationKeys.Lists.Polls.Fields.MeetingOrganizer];
        const meetingAttendeesFieldName = this.config[ConfigurationKeys.Lists.Polls.Fields.MeetingAttendees];
        const templateFieldName = this.config[ConfigurationKeys.Lists.Polls.Fields.Template];
        const startDateFieldName = this.config[ConfigurationKeys.Lists.Polls.Fields.StartDate];
        const endDateFieldName = this.config[ConfigurationKeys.Lists.Polls.Fields.EndDate];
        const meetingNameFieldName = this.config[ConfigurationKeys.Lists.Polls.Fields.MeetingName];

        const fields = {};
        fields[meetingIdFieldName] = meetingId;
        fields[meetingOrganizerFieldName] = meeting.organizer?.upn;
        fields[meetingAttendeesFieldName] = meeting.attendees?.map(a => a.upn).join('; ');
        fields[templateFieldName] = templateId;
        fields[startDateFieldName] = meeting.startDateTime?.toISOString();
        fields[endDateFieldName] = meeting.endDateTime?.toISOString();
        fields[meetingNameFieldName] = meeting.subject;

        const listItem = {
            fields: fields
        };

        const result = await this.postData(url, listItem);

        const poll = await this.getMeetingPoll(meetingId);

        return poll;
    }

    public getQuestionsList = async (templateId: string): Promise<Array<IQuestion>> => {

        const template = await this.getTemplateById(templateId);

        if (!template || !template.questionsIds || template.questionsIds.length < 1) return new Promise(resolve => resolve([]));

        const titleFieldName = this.config[ConfigurationKeys.Lists.Questions.Fields.Title];
        const typeFieldName = this.config[ConfigurationKeys.Lists.Questions.Fields.Type];
        const isRequiredFieldName = this.config[ConfigurationKeys.Lists.Questions.Fields.IsRequired];

        const listTitle = this.config[ConfigurationKeys.Lists.Questions.Title];
        const selectedFields: Array<string> = compact([titleFieldName, typeFieldName, isRequiredFieldName]);

        const batchData = {
            "requests": template.questionsIds.map((id, index) => {
                return {
                    "id": index + 1,
                    "method": "GET",
                    "url": `/sites/${this.siteDomain}:${this.serverRelativeUrl}:/lists/${listTitle}/items/${id}?expand=fields(select=${selectedFields.join(',')})`
                }
            })
        }

        const url = "https://graph.microsoft.com/v1.0/$batch";

        const result = await this.postData(url, batchData);
        if (!result || !result.responses) return [];

        const sortedResonses = result.responses.sort((a, b) => +a.id - +b.id);
        const questions: Array<IQuestion> = compact(sortedResonses.map(response => {
            if (!response || response.status !== 200 || !response.body) return null;
            return {
                id: response.body["id"],
                title: response.body.fields[titleFieldName],
                type: this.getQuestionType(response.body.fields[typeFieldName]),
                isRequired: response.body.fields[isRequiredFieldName]
            } as IQuestion;

        }));
        return questions;
    }

    public getResponsesForPoll = async (pollId: string, forCurrentUserOnly: boolean = true): Promise<Array<IResponse>> => {

        if (!this.context.userObjectId || !pollId) return new Promise(resolve => resolve([]));

        const userIdFieldName = this.config[ConfigurationKeys.Lists.Responses.Fields.UserId];
        const meetingIdFieldName = this.config[ConfigurationKeys.Lists.Responses.Fields.MeetingId];
        const pollIdFieldName = this.config[ConfigurationKeys.Lists.Responses.Fields.PollId];
        const questionIdFieldName = this.config[ConfigurationKeys.Lists.Responses.Fields.QuestionId];
        const responseFieldName = this.config[ConfigurationKeys.Lists.Responses.Fields.Response];

        const listTitle = this.config[ConfigurationKeys.Lists.Responses.Title];
        const selectedFields: Array<string> = compact([userIdFieldName, meetingIdFieldName, pollIdFieldName, questionIdFieldName, responseFieldName]);

        let url = `https://graph.microsoft.com/v1.0/sites/${this.siteDomain}:${this.serverRelativeUrl}:/lists/${listTitle}/items?expand=fields(select=${selectedFields.join(',')})&$filter=fields/${pollIdFieldName} eq '${pollId}'`;
        if (forCurrentUserOnly) {
            url += ` and fields/${userIdFieldName} eq '${this.context.userObjectId}'`;
        }

        const result = await this.getData(url, true);
        if (!result || !result.value) return [];
        const responses: Array<IResponse> = compact(result.value.map(item => {
            if (!item || !item.fields) return null;
            return {
                id: item["id"],
                userId: item.fields[userIdFieldName],
                meetingId: item.fields[meetingIdFieldName],
                pollId: item.fields[pollIdFieldName],
                questionId: item.fields[questionIdFieldName],
                response: item.fields[responseFieldName],
            } as IResponse;
        }));
        return responses;
    }

    public postQuestionsResponses = async (meetingId: string, pollId: string, newResponses: Array<IResponse>, existingResponses: Array<IResponse>) => {
        if (!meetingId) throw new Error("GraphService: postQuestionResponse - Empty meetingId");
        if (!pollId) throw new Error("GraphService: postQuestionResponse - Empty pollId");
        if (!newResponses) throw new Error("GraphService: postQuestionResponse - Empty newResponses");
        if (!existingResponses) throw new Error("GraphService: postQuestionResponse - Empty existingResponses");

        const userIdFieldName = this.config[ConfigurationKeys.Lists.Responses.Fields.UserId];
        const meetingIdFieldName = this.config[ConfigurationKeys.Lists.Responses.Fields.MeetingId];
        const tenantIdFieldName = this.config[ConfigurationKeys.Lists.Responses.Fields.TenantId];
        const pollIdFieldName = this.config[ConfigurationKeys.Lists.Responses.Fields.PollId];
        const questionIdFieldName = this.config[ConfigurationKeys.Lists.Responses.Fields.QuestionId];
        const responseFieldName = this.config[ConfigurationKeys.Lists.Responses.Fields.Response];

        const listTitle = this.config[ConfigurationKeys.Lists.Responses.Title];

        const requests = [];
        requests.push(...newResponses.map((response, index) => {

            const fields = {};
            fields[userIdFieldName] = this.context.userObjectId;
            fields[meetingIdFieldName] = meetingId;
            fields[tenantIdFieldName] = this.context.tid;
            fields[questionIdFieldName] = response.questionId;
            fields[responseFieldName] = response.response;
            fields[pollIdFieldName] = pollId;

            const listItem = {
                fields: fields
            };

            return {
                "id": index + 1,
                "method": "POST",
                "url": `/sites/${this.siteDomain}:${this.serverRelativeUrl}:/lists/${listTitle}/items`,
                "headers": {
                    'Content-Type': 'application/json'
                },
                "body": listItem
            }
        }));

        requests.push(...existingResponses.map((response, index) => {
            const fields = {};
            fields[userIdFieldName] = this.context.userObjectId;
            fields[meetingIdFieldName] = meetingId;
            fields[tenantIdFieldName] = this.context.tid;
            fields[questionIdFieldName] = response.questionId;
            fields[responseFieldName] = response.response;
            fields[pollIdFieldName] = pollId;

            const listItem = {
                fields: fields
            };

            return {
                "id": newResponses.length + index + 1,
                "method": "PATCH",
                "url": `/sites/${this.siteDomain}:${this.serverRelativeUrl}:/lists/${listTitle}/items/${response.id}`,
                "headers": {
                    'Content-Type': 'application/json'
                },
                "body": listItem
            }
        }));

        const batchData = {
            "requests": requests
        };

        const url = "https://graph.microsoft.com/v1.0/$batch";

        const result = await this.postData(url, batchData);
    }

    public getMeetingTabId = async (tabName: string): Promise<string> => {

        const url = `https://graph.microsoft.com/v1.0/chats/${this.context.chatId}/tabs`;

        const result = await this.getData(url, true);
        if (!result || !result.value || !result.value[0]) return null;
        const { value } = result;

        const meetingTab = value.filter(t => t["displayName"]?.startsWith(tabName))[0];
        return !!meetingTab && !!meetingTab.id ? meetingTab.id : null;
    }

    public getUsersInfo = async (userIds: Array<string>): Promise<Array<IUser>> => {
        if (!userIds || userIds.length < 1) return new Promise(resolve => resolve([]));

        const selectedFields: Array<string> = compact([
            "id",
            "displayName"
        ]);

        const batchData = {
            "requests": userIds.map((userId, index) => {
                return {
                    "id": index + 1,
                    "method": "GET",
                    "url": `/users/${userId}?$select=${selectedFields.join(",")}`
                }
            })
        }

        const url = "https://graph.microsoft.com/v1.0/$batch";

        const result = await this.postData(url, batchData);
        if (!result || !result.responses) return [];

        const users: Array<IUser> = compact(result.responses.map(response => {
            if (!response || response.status !== 200 || !response.body) return null;
            return {
                id: response.body["id"],
                displayName: response.body["displayName"],
            } as IUser;

        }));
        return users;
    }

    private getTemplateById = async (templateId): Promise<ITemplate> => {

        const titleFieldName = this.config[ConfigurationKeys.Lists.Templates.Fields.Title];
        const descriptionFieldName = this.config[ConfigurationKeys.Lists.Templates.Fields.Description];
        const questionsFieldName = this.config[ConfigurationKeys.Lists.Templates.Fields.Questions];

        const listTitle = this.config[ConfigurationKeys.Lists.Templates.Title];
        const selectedFields: Array<string> = compact([titleFieldName, descriptionFieldName, questionsFieldName]);
        const url = `https://graph.microsoft.com/v1.0/sites/${this.siteDomain}:${this.serverRelativeUrl}:/lists/${listTitle}/items/${templateId}?expand=fields(select=${selectedFields.join(',')})`;

        const result = await this.getData(url);
        if (!result) return null;
        const template: ITemplate = {
            id: result["id"],
            title: result.fields[titleFieldName],
            description: result.fields[descriptionFieldName],
            questionsIds: !!result.fields[questionsFieldName]
                ? result.fields[questionsFieldName].map(i => i["LookupId"])
                : []
        };
        return template;
    }

    public postChatMessage = async (tabId: string) => {

        const url = `https://graph.microsoft.com/v1.0/chats/${this.context.chatId}/messages`;


        const id = GetGuid();

        // if tabId is not found return a message without a link
        const link = !!tabId ?
            `<a href='https://teams.microsoft.com/_#/tab::${tabId}/${this.context.chatId}?ctx=chat'>Meetings Survey</a>`
            : "Meetings Survey";
        const data = {
            "body": {
                "contentType": "html",
                "content": `<attachment id="${id}"></attachment>`
            },
            "attachments": [
                {
                    "id": id,
                    "contentType": "application/vnd.microsoft.card.thumbnail",
                    "contentUrl": null,
                    "content": `{ 
                        "title": "Meetings Survey", "text": "The survey was launched for this meeting. Go to ${link} tab",}`,
                    "name": null,
                    "thumbnailUrl": null
                }
            ]
        };

        const result = await this.postData(url, data);

        return null;
    }

    private getData = async (url: string, needPreferHeader: boolean = false) => {
        try {
            const requestHeaders: HeadersInit = new Headers();

            requestHeaders.set('Accept', 'application/json');
            requestHeaders.set('Authorization', "bearer " + this.aadToken);

            if (needPreferHeader) {
                requestHeaders.set('Prefer', 'HonorNonIndexedQueriesWarningMayFailRandomly');
            }

            const response = await fetch(url,
                {
                    method: 'GET',
                    headers: requestHeaders,
                    mode: 'cors',
                    cache: 'default'
                });

            if (!response) return null;

            if (!response.ok) {
                const errorResponseJson = await response.json();
                const message = errorResponseJson?.error?.message;
                throw new Error(message);
            }

            const responseJson = await response.json();
            if (!responseJson) return null;
            return responseJson;
        } catch (error) {
            const message = `Error in GraphService - Url: ${url}, Error: ${error}`;
            console.error(message);
            throw new Error(message);
        }
    }

    private postData = async (url: string, data) => {
        try {
            const requestHeaders: HeadersInit = new Headers();
            requestHeaders.set('Content-Type', 'application/json');
            requestHeaders.set('Accept', 'application/json');
            requestHeaders.set('Authorization', "bearer " + this.aadToken);

            const response = await fetch(url,
                {
                    method: 'POST',
                    headers: requestHeaders,
                    mode: 'cors',
                    cache: 'default',
                    body: JSON.stringify(data)
                });
            if (!response) return null;

            if (!response.ok) {
                const errorResponseJson = await response.json();
                const message = errorResponseJson?.error?.message;
                throw new Error(message);
            }

            const responseJson = await response.json();
            if (!responseJson) return null;
            return responseJson;
        } catch (error) {
            const message = `Error in GraphService - Url: ${url}, Error: ${error}`;
            console.error(message);
            throw new Error(message);
        }
    }

    private initUrls = () => {
        const siteUrl = this.config[ConfigurationKeys.SourceSiteUrl];
        if (!siteUrl || (!siteUrl.startsWith("http://") && !siteUrl.startsWith("https://"))) throw new Error("Invalid siteUrl");
        const siteUrlSplitted = siteUrl.split("/sites/");
        this.siteDomain = siteUrlSplitted[0].split("://")[1];
        this.serverRelativeUrl = siteUrlSplitted.length > 1
            ? `/sites/${siteUrlSplitted[1]}`
            : "/";
    }

    private getQuestionType = (type: string): QuestionType => {
        switch (type) {
            case "Yes/No":
                return QuestionType.YesNo;
            case "Text":
            default:
                return QuestionType.Text;
        }
    }
}