import { ConfigurationKeys } from "../constants";

export default class ConfigurationService {

    public static getConfig = async (): Promise<object> => {

        const serverURL = `${window.location.origin}/getConfig`;
        const response = await fetch(serverURL);
        if (response) {
            const data = await response.json();

            // add validation here if needed
            const sourceSiteUrl = data[ConfigurationKeys.SourceSiteUrl];
            data[ConfigurationKeys.SourceSiteUrl] = !!sourceSiteUrl ? sourceSiteUrl.replace(/\/+$/, "") : ""

            return data;
        }
        return {};
    }
}