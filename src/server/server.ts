import * as Express from "express";
import * as http from "http";
import * as path from "path";
import * as morgan from "morgan";
import { MsTeamsApiRouter, MsTeamsPageRouter } from "express-msteams-host";
import * as debug from "debug";
import * as compression from "compression";
import jwt_decode, { JwtPayload } from 'jwt-decode';
import fetch from 'node-fetch';

// Initialize debug logging module
const log = debug("msteams");

log("Initializing Microsoft Teams Express hosted App...");

// The import of components has to be done AFTER the dotenv config
// eslint-disable-next-line import/first
import * as allComponents from "./TeamsAppsComponents";

// Create the Express webserver
const express = Express();
const port = process.env.port || process.env.PORT || 3007;

// Inject the raw request body onto the request object
express.use(Express.json({
    verify: (req, res, buf: Buffer, encoding: string): void => {
        (req as any).rawBody = buf.toString();
    }
}));
express.use(Express.urlencoded({ extended: true }));

// Express configuration
express.set("views", path.join(__dirname, "/"));

// Add simple logging
express.use(morgan("tiny"));

// Add compression - uncomment to remove compression
express.use(compression());

// Add /scripts and /assets as static folders
express.use("/scripts", Express.static(path.join(__dirname, "web/scripts")));
express.use("/assets", Express.static(path.join(__dirname, "web/assets")));

// routing for bots, connectors and incoming web hooks - based on the decorators
// For more information see: https://www.npmjs.com/package/express-msteams-host
express.use(MsTeamsApiRouter(allComponents));

// routing for pages for tabs and connector configuration
// For more information see: https://www.npmjs.com/package/express-msteams-host
express.use(MsTeamsPageRouter({
    root: path.join(__dirname, "web/"),
    components: allComponents
}));

// Set default web page
express.use("/", Express.static(path.join(__dirname, "web/"), {
    index: "index.html"
}));

// get config from environment variables
express.get('/getConfig', async (req, res) => {
    const configPrefix = "MEETINGSSURVEY_";
    const config = {};    
    const envObjectKeys = Object.keys(process.env);
    envObjectKeys.forEach(key=>{
        if(!!key && key.startsWith(configPrefix) && key !== "MEETINGSSURVEY_APP_SECRET"){
            config[key]=process.env[key];
        }
    })

    return  res.send(config);
});

// exchange id token to aad token
express.get('/getGraphAccessToken', async (req, res) => {

    const ssoToken = req.query.ssoToken as string;
    let tenantId = jwt_decode<JwtPayload>(ssoToken)['tid']; //Get the tenant ID from the decoded token
    let accessTokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    //Create your access token query parameters
    //Learn more: https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow#first-case-access-token-request-with-a-shared-secret
    const client_id: string = process.env.MEETINGSSURVEY_APP_ID as string;
    const client_secret: string = process.env.MEETINGSSURVEY_APP_SECRET as string;

    if (!client_id) {
        console.log("getGraphAccessToken: Client Id is not valid");
        res.status(400).json({ error: 'Client Id is not valid' });
    }

    if (!client_secret) {
        console.log("getGraphAccessToken: Client Secret is not valid");
        res.status(400).json({ error: 'Client Secret is not valid' });
    }

    const getGraphScope = (): string => {
        const scopes = ["ChatMessage.Send", "OnlineMeetings.Read", "Sites.ReadWrite.All", "TeamsTab.Read.All", "User.Read.All"];
        return scopes.map(scope => `https://graph.microsoft.com/${scope}`).join(" ");
    }

    let accessTokenQueryParams = {
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        client_id: client_id,
        client_secret: client_secret,
        assertion: req.query.ssoToken as string,
        scope: getGraphScope(),
        requested_token_use: "on_behalf_of",
    };

    let body = new URLSearchParams(accessTokenQueryParams).toString();

    let accessTokenReqOptions = {
        method: 'POST',
        headers: {
            Accept: "application/json",
            "Content-Type": "application/x-www-form-urlencoded"
        },
        body: body
    };

    let response = await fetch(accessTokenEndpoint, accessTokenReqOptions).catch(handleQueryError) as any;

    let data = await response.json();
    console.log(`${data.token_type} token received`);
    if (!response.ok) {
        if ((data.error === 'invalid_grant') || (data.error === 'interaction_required')) {
            //This is expected if it's the user's first time running the app ( user must consent ) or the admin requires MFA
            console.log("User must consent or perform MFA. You may also encouter this error if your client ID or secret is incorrect.");
            res.status(403).json({ error: 'consent_required' }); //This error triggers the consent flow in the client.
        } else {
            //Unknown error
            console.log('Could not exchange access token for unknown reasons.');
            res.status(500).json({ error: 'Could not exchange access token' });
        }
    } else {
        //The on behalf of token exchange worked. Return the token (data object) to the client.
        res.send(data);
    }
});

let handleQueryError = function (err: string) {
    console.log("handleQueryError called: ", err);
};

// Set the port
express.set("port", port);

// Start the webserver
http.createServer(express).listen(port, () => {
    log(`Server running on ${port}`);
});
