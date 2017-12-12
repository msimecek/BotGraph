require("dotenv").config();

import * as restify from "restify"
import * as builder from "botbuilder";
import * as path from "path";

import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import * as MicrosoftGraphClient from "@microsoft/microsoft-graph-client";

import * as expressSession from 'express-session';
import { AuthHelper } from "./helpers/AuthHelper";

var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3979, function () {
    console.log("%s listening to %s", server.name, server.url);
});

// -- bot code --
var connector = new builder.ChatConnector({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});

server.post("api/messages", connector.listen());
server.get("/code", restify.plugins.serveStatic({
    directory: path.join(__dirname, "public"),
    file: "code.html"
}));

var bot = new builder.UniversalBot(connector);

//--auth setup --

server.use(restify.plugins.queryParser());
server.use(restify.plugins.bodyParser());
server.use(expressSession({ secret: process.env.BOTAUTH_SECRET, resave: true, saveUninitialized: false }));

const authHelper = new AuthHelper(server, bot);

//--- bot code --

bot.dialog("/", [].concat(
    (session, args, next) => {
        session.send("Hello, world.");
        next();
    },
    authHelper.getAccessToken(),
    (session, results, next) => {
        if (results.response !== null) {
            var client = MicrosoftGraphClient.Client.init({
                authProvider: (done) => {
                    done(null, results.response);
                }
            });

            var messages: MicrosoftGraph.Message[];
            client.api("me/messages").top(5).select("subject").get().then((res) => {
                messages = res.value;
                session.send(JSON.stringify(messages));
            })
        }
    })
);

bot.dialog("/logout", (session) => {
    authHelper.logout(session);
    session.send("Logged out.");
    session.endDialog();
}).triggerAction({ matches: /logout/ });