export class ConfigurationKeys {

    static SourceSiteUrl: string = "MEETINGSSURVEY_SOURCE_SITE_URL";

    static AppRegistration = class AppRegistration {
        static AppId: string = "MEETINGSSURVEY_APP_ID";
        static AppUri: string = "MEETINGSSURVEY_APP_URI";
        static AppSecret: string = "MEETINGSSURVEY_APP_SECRET";
    };

    static Lists = class Lists {
        static Questions = class Questions {
            static Title: string = "MEETINGSSURVEY_QUESTIONS_LIST_TITLE";
            static Fields = class Fields {
                static Title: string = "MEETINGSSURVEY_QUESTIONS_LIST_TITLE_FIELD_NAME";
                static Type: string = "MEETINGSSURVEY_QUESTIONS_LIST_TYPE_FIELD_NAME";
                static IsRequired: string = "MEETINGSSURVEY_QUESTIONS_LIST_IS_REQUIRED_FIELD_NAME";
            };
        };

        static Templates = class Questions {
            static Title: string = "MEETINGSSURVEY_TEMPLATES_LIST_TITLE";
            static Fields = class Fields {
                static Title: string = "MEETINGSSURVEY_TEMPLATES_LIST_TITLE_FIELD_NAME";
                static Description: string = "MEETINGSSURVEY_TEMPLATES_LIST_DESCRIPTION_FIELD_NAME";
                static Questions: string = "MEETINGSSURVEY_TEMPLATES_LIST_QUESTIONS_FIELD_NAME";
            };
        };

        static Polls = class Questions {
            static Title: string = "MEETINGSSURVEY_POLLS_LIST_TITLE";
            static Fields = class Fields {
                static MeetingId: string = "MEETINGSSURVEY_POLLS_LIST_MEETING_ID_FIELD_NAME";
                static MeetingOrganizer: string = "MEETINGSSURVEY_POLLS_LIST_MEETING_ORGANIZER_FIELD_NAME";
                static MeetingAttendees: string = "MEETINGSSURVEY_POLLS_LIST_MEETING_ATTENDEES_FIELD_NAME";
                static Template: string = "MEETINGSSURVEY_POLLS_LIST_POLL_TEMPLATE_FIELD_NAME";
                static StartDate: string = "MEETINGSSURVEY_POLLS_LIST_MEETING_START_DATE_FIELD_NAME";
                static EndDate: string = "MEETINGSSURVEY_POLLS_LIST_MEETING_END_DATE_FIELD_NAME";
                static MeetingName: string = "MEETINGSSURVEY_POLLS_LIST_MEETING_NAME_FIELD_NAME";
            };
        };

        static Responses = class Questions {
            static Title: string = "MEETINGSSURVEY_RESPONSES_LIST_TITLE";
            static Fields = class Fields {
                static UserId: string = "MEETINGSSURVEY_RESPONSES_LIST_USER_ID_FIELD_NAME";
                static MeetingId: string = "MEETINGSSURVEY_RESPONSES_LIST_MEETING_ID_FIELD_NAME";
                static TenantId: string = "MEETINGSSURVEY_RESPONSES_LIST_TENANT_ID_FIELD_NAME";
                static QuestionId: string = "MEETINGSSURVEY_RESPONSES_LIST_QUESTION_ID_FIELD_NAME";
                static Response: string = "MEETINGSSURVEY_RESPONSES_LIST_RESPONSE_FIELD_NAME";
                static PollId: string = "MEETINGSSURVEY_RESPONSES_LIST_POLL_ID_FIELD_NAME";
            };
        };
    };
}

