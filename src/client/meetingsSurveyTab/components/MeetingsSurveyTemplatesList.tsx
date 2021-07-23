
import { Button, Loader } from "@fluentui/react-northstar";
import * as React from "react";
import { ITemplate } from "../interfaces";
import GraphService from "../services/GraphService";
import { GetExceptionMessage } from "../utils/MessageUtils";
import '../styles/MeetingsSurveyContainer.scss';
import strings from "../loc/localizedStrings";

export interface IMeetingsSurveyTemplatesListProps {
    context: microsoftTeams.Context;
    graphService: GraphService;
    onTemplateLaunched: (templateId: string) => void;
    onGetError: (message: string) => void;
}

export interface IMeetingsSurveyTemplatesListState {
    isLoading: boolean;
    templates: Array<ITemplate>;
}

export class MeetingsSurveyTemplatesList extends React.Component<IMeetingsSurveyTemplatesListProps, IMeetingsSurveyTemplatesListState> {

    constructor(props) {
        super(props);
        this.state = {
            isLoading: true,
            templates: []
        }

        this.handlers.loadTemplatesInformation();
    }

    public render(): React.ReactElement<IMeetingsSurveyTemplatesListProps> {
        const { isLoading } = this.state;
        const controls = [];
        if (isLoading) {
            controls.push(this._render.loader());
        } else {
            controls.push(
                this._render.content()
            );
        }

        return <div>
            {controls}
        </div>;
    }

    private _render = {
        loader: (): JSX.Element => {
            return <Loader label={strings.loadingTemplatesResultsMessage} />;
        },
        content: (): JSX.Element => {
            const { templates } = this.state;
            if (!templates || templates.length < 1) return <div>{strings.templatesNotFoundMessage}</div>;

            return <div className="meetingSurveyTemplatesListContainer">
                <div className="meetingSurveyTemplatesHeader">
                    {strings.selectTemplateHeader}
                </div>
                <div className="meetingSurveyTemplatesContainer">
                    {templates.map(this._render.template)}
                </div>
            </div>
        },
        template: (template: ITemplate): JSX.Element => {
            const { onTemplateLaunched } = this.props;

            return <div
                key={template.id}
                className="meetingSurveyTemplateContainer"
            >
                <div title={template.title} className="meetingSurveyTemplateTitle">{template.title}</div>
                <div title={template.description} className="meetingSurveyTemplateDescription">{template.description}</div>
                <div className="meetingSurveyTemplateActions">
                    <Button
                        primary
                        content={strings.launchButtonLabel}
                        onClick={() => onTemplateLaunched(template.id)}
                    />
                </div>
            </div>
        }
    };

    private handlers = {
        loadTemplatesInformation: async () => {
            const { graphService, onGetError } = this.props;

            try {

                // load templates list
                const templatesList = await graphService.getTemplatesList();

                this.setState({
                    isLoading: false,
                    templates: templatesList
                });

            } catch (error) {
                this.setState({
                    isLoading: false
                });
                onGetError(GetExceptionMessage(error));
            }
        }
    }
}