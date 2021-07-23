# Meetings Survey app configuration

## 1. Create lists
Create or choose existing SharePoint site collection that will be used as data storage for Meetings Survey app. Note that users that are going to use this app are expected to have at least members permissions (they need to be able to create and edit list items at this site). You need to create the folliwing lists manually (note that you can use different list and field names but then you will need to update them in application variables):
### 1.1. Questions list
This list contains questions and their options (such as type). The purpose of storing them separately is sharing same questions between different templates.
- List title: **QVCQuestions**
- List fields:
    - **Title** - Single line of text (you can use existing field)
    - **QuestionType** - Choice: *Text, Yes/No* options
    - **IsRequired** - Boolean: *Yes* by default
 
### 1.2. Templates list
This list represents survey templates that are basically a set of questions.
- List title: **QVCTemplates**
- List fields:
    - **Title** - Single line of text (you can use existing field)
     - **TemplateDescription** - Multiline text
     - **TemplateQuestions** - Multi-lookup to *QVCQuestions* *Title* column
### 1.3. Poll list
This list represents actual polls that were launched for particular meetings. The item in this list is created when meeting organizer clicks the Launch button. It contains information about the meeting and id of a template.
- List title: **QVCActivePolls**
- List fields:
     - **MeetingId** - Single line of text
     - **MeetingOrganizer** - Single line of text
     - **MeetingAttendees** - Multiline text
     - **PollTemplate** - Single line of text
     - **MeetingStartDate** - **Single line of text**
     - **MeetingEndDate** - **Single line of text**
     - **MeetingName** - Single line of text
     Make *MeetingId* column indexed (via list settings) because it will be used in queries.
### 1.4. Responses list
This list represents user responses for questions.Each items contains a response of a particular user to a particular question. It is expected that this list will contain a big number of items if the app will be used extensively.
- List title: **QVCPollsResponses**
- List fields:
     - **UserId** - Single line of text
     - **MeetingId** - Single line of text
     - **TenantId** - Single line of text
     - **QuestionId** - Single line of text
     - **PollId** - Single line of text
     - **Response** - Multiline text
Make *PollId* and *UserId* columns indexed (via list settings) because they will be used in queries.
## 2. Create App Registration
You need to create App Registration to get permissions for Graph API. Here is general information about this:
https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso
1. Go to https://portal.azure.com
2. Login with account that has Azure subscription
3. Go to Azure Active Directory -> App Registrations -> New registration
4. Give your App Registration a name and click Register. Copy your App ID.
5. Go to Expose and API -> click Set -> Specify your url in the following format: **api://<your-app-domain>/<app-id>**, where <your-app-domain> is the domain name of your web application (should be like *teams-meetings-sample-app.azurewebsites.net*) and <app-id> is client id of your app registration. Click Save. Copy your App URI.
6. Click Add a scope. Enter *access_as_user* as the Scope name. Select *Admins and Users* in Who can consent toggle. Enter admin and user consent titles and descriptions. This will be displayed in the consent window. Makw sure that *State* is set to *Enabled*. Click *Add scope*.
7. In the Authorized client applications section, identify the applications that you want to authorize for your app’s web application. Select Add a client application. Enter each of the following client IDs and select the authorized scope you created in the previous step:
    - 1fec8e78-bce4-4aaf-ab1b-5451cc387264 - for Teams mobile or desktop application.
    - 5e3ce6c0-2b1f-4285-8d4b-75ee78787346 - for Teams web application.
8. Navigate to API Permissions. Select Add a permission -> Microsoft Graph -> Delegated permissions, then add the following permissions from Graph API:
    - email
    - offline_access
    - OpenId
    - profile
    - ChatMessage.Send
    - OnlineMeetings.Read
    - Sites.ReadWrite.All
    - TeamsTab.Read.All
    - User.Read.All
9. Click *Grant Admin consent for <tenant_name>*
10. Navigate to *Authentication*. Click Add a platform -> Web -> Enter redirect Uri in format: **https://<your-app-domain>/meetingsSurveyTab/auth-end.html**. Check *ID token* and *Access token* boxes. Click *Save*.
11. Go to *Certificates & Secrets*. Click *New client secret*. Give it a name and select expiration date (note that you will have to update it after it expires). Save App Secret value.

## 3. Deploy application to Azure
To get general information about deploying process please refer to a link.
https://docs.microsoft.com/en-us/azure/app-service/quickstart-nodejs
- Install Azure App Service extension to Visual Studio Code.
- Sign in to Azure from this extension
- Click on the up arrow icon in order to deploy your project. The exact deployment process will depend on the platform selected ane service plan.
- After the app is deployed make sure that it is running. Then go to app settings page in Azure and click *Configuration* in left hand panel.
- You need to edit *Application settings* that are going to be environment variables for an app. Click *Advanced edit*. This allows you to paste settings in JSON format. Paste this section:

          {
            "name": "MEETINGSSURVEY_APP_ID",
            "value": "<app-registration-id>",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_APP_SECRET",
            "value": "<app-registration-secret>",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_APP_URI",
            "value": "<app-registration-uri>",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_SOURCE_SITE_URL",
            "value": "<url-of-source-sharepoint-site-collection>",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_QUESTIONS_LIST_TITLE",
            "value": "QVCQuestions",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_QUESTIONS_LIST_TITLE_FIELD_NAME",
            "value": "Title",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_QUESTIONS_LIST_TYPE_FIELD_NAME",
            "value": "QuestionType",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_QUESTIONS_LIST_IS_REQUIRED_FIELD_NAME",
            "value": "IsRequired",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_TEMPLATES_LIST_TITLE",
            "value": "QVCTemplates",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_TEMPLATES_LIST_TITLE_FIELD_NAME",
            "value": "Title",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_TEMPLATES_LIST_DESCRIPTION_FIELD_NAME",
            "value": "TemplateDescription",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_TEMPLATES_LIST_QUESTIONS_FIELD_NAME",
            "value": "TemplateQuestions",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_POLLS_LIST_TITLE",
            "value": "QVCActivePolls",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_POLLS_LIST_POLL_TEMPLATE_FIELD_NAME",
            "value": "PollTemplate",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_POLLS_LIST_MEETING_ID_FIELD_NAME",
            "value": "MeetingId",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_POLLS_LIST_MEETING_NAME_FIELD_NAME",
            "value": "MeetingName",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_POLLS_LIST_MEETING_ORGANIZER_FIELD_NAME",
            "value": "MeetingOrganizer",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_POLLS_LIST_MEETING_ATTENDEES_FIELD_NAME",
            "value": "MeetingAttendees",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_POLLS_LIST_MEETING_START_DATE_FIELD_NAME",
            "value": "MeetingStartDate",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_POLLS_LIST_MEETING_END_DATE_FIELD_NAME",
            "value": "MeetingEndDate",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_RESPONSES_LIST_TITLE",
            "value": "QVCPollsResponses",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_RESPONSES_LIST_MEETING_ID_FIELD_NAME",
            "value": "MeetingId",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_RESPONSES_LIST_POLL_ID_FIELD_NAME",
            "value": "PollId",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_RESPONSES_LIST_QUESTION_ID_FIELD_NAME",
            "value": "QuestionId",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_RESPONSES_LIST_RESPONSE_FIELD_NAME",
            "value": "Response",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_RESPONSES_LIST_TENANT_ID_FIELD_NAME",
            "value": "TenantId",
            "slotSetting": false
          },
          {
            "name": "MEETINGSSURVEY_RESPONSES_LIST_USER_ID_FIELD_NAME",
            "value": "UserId",
            "slotSetting": false
          }
          
    Make sure that you replaced all placeholders with actual values. List titles and field names are given by default, so if you made changes in Section 1 make corresponding adjustments.
    
- Go to Authentication section. Enable App Service Authentication. Select *Log in with Azure Active Directory*. Select previously created App Registration from dropdown or enter its parameters if autoselect in not available. Then go to App Registration, open *Authentication* section and add the following Redirect Uri:
`
https://<your-app-domain>/.auth/login/aad/callback
`
Click *Save*.

## 4. Upload app package to Teams
- Go to *manifest.json* file in your project.
- Generate unique guid for your app id.
- Replace <your-app-domain> and <app-registration-id> with real values
- Gp to ./src/manifest folder and create .zip file from its content
- Go to Teams -> More apps -> Upload a custom app -> Upload for my org -> Select .zip package
- The app is uploaded in your Teams. You can add it as a tab from Meeting settings
