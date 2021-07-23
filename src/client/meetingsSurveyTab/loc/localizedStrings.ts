import LocalizedStrings from 'react-localization';

const localizedStrings = new LocalizedStrings({
    en: {
        questionPlaceholder: "Enter answer here...",
        submitButtonLabel: "Submit",
        closeButtonLabel: "Close",
        responsesDetailsHeader: "Responses details",
        viewDetailsButtonLabel: "View Details",
        resfreshButtonLabel: "Refresh",
        launchButtonLabel: "Launch",
        selectTemplateHeader: "Select template to launch:",
        surveyHeader: "Survey",
        emptyPollLabel: "There are no available surveys at this time",

        // messages
        configurationMessage: "Click 'Save' to add app to a meeting",
        removeTabMessage: "You're about to remove your tab...",
        removeTabAdditionalMessage: "Recorded details of the poll will be stored in database",
        redirectToConsentPageMessage: "Redirecting to consent page...",
        consentFlowCompletedMessage: "Consent flow completed",
        teamsContextOnlyMessage: "This app is only available from Teams context",

        // loading messages
        loadingAuthenticationTokenMessage: "Loading authentication token...",
        loadingMeetngInfoMessage: "Loading meeting info...",
        settingTemplateMessage: "Setting up the template...",
        loadingPollMessage: "Loading poll...",
        loadingResultsMessage: "Loading results...",
        loadingTemplatesResultsMessage: "Loading templates...",

        // error messages
        aadTokenNotValidMessage: "aadToken is not valid",
        ssoErrorMessage: "An SSO error occurred",
        consentFailedMessage: "Consent failed",
        pollNotCreatedMessage: "Poll wasn't created",
        pollExpirationMessage: "The poll cannot be submitted after expiration time",
        questionsNotFoundMessage: "No questions were found for this poll",
        templatesNotFoundMessage: "Poll templates were not found"

    }
});

export default localizedStrings;