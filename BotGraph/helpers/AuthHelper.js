"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
require('dotenv').config();
//import * as https from 'https';
const request = require("request");
const builder = require("botbuilder");
const botauth = require("botauth");
const passport_azure_ad_1 = require("passport-azure-ad");
class AuthHelper {
    constructor(server, bot) {
        this._botAuth = new botauth.BotAuthenticator(server, bot, {
            session: true,
            baseUrl: `https://${process.env.WEBSITE_HOSTNAME}`,
            secret: process.env.BOTAUTH_SECRET,
            successRedirect: '/code'
        });
        this._botAuth.provider('aadv2', (options) => {
            return new passport_azure_ad_1.OIDCStrategy({
                redirectUrl: `https://${process.env.WEBSITE_HOSTNAME}/botauth/aadv2/callback`,
                realm: 'common',
                clientID: process.env.CLIENT_ID,
                clientSecret: process.env.CLIENT_SECRET,
                identityMetadata: 'https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration',
                skipUserProfile: false,
                validateIssuer: false,
                allowHttpForRedirectUrl: true,
                responseType: 'code',
                responseMode: 'query',
                scope: ['profile', 'offline_access', 'calendars.readwrite', 'files.readwrite', 'mail.read'],
                passReqToCallback: true
            }, (req, iss, sub, profile, accessToken, refreshToken, done) => {
                if (!profile.displayName) {
                    return done(new Error('No oid found'), null);
                }
                profile.accessToken = accessToken;
                profile.refreshToken = refreshToken;
                done(null, profile);
            });
        });
    }
    getAccessToken() {
        return [].concat(this._botAuth.authenticate('aadv2', {}), (session, results, next) => {
            // get token with refresh token
            let user = this._botAuth.profile(session, 'aadv2');
            this.getAccessTokenWithRefreshToken(user.refreshToken).then((data) => {
                session.userData.accessToken = data.accessToken;
                session.userData.refreshToken = data.refreshToken;
                next({ response: data.accessToken, resumed: builder.ResumeReason.forward });
            }, (err) => {
                next({ response: null, resumed: builder.ResumeReason.forward });
            });
        });
    }
    getAccessTokenWithRefreshToken(refreshToken) {
        return new Promise((resolve, reject) => {
            var data = 'grant_type=refresh_token'
                + '&refresh_token=' + refreshToken
                + '&client_id=' + process.env.CLIENT_ID
                + '&client_secret=' + encodeURIComponent(process.env.CLIENT_SECRET);
            var options = {
                method: 'POST',
                url: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
                body: data,
                json: true,
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
            };
            request(options, function (err, res, body) {
                if (err)
                    return reject(err);
                resolve({
                    accessToken: body.access_token,
                    refreshToken: body.refresh_token
                });
            });
        });
    }
    logout(session) {
        this._botAuth.logout(session, "aadv2");
    }
}
exports.AuthHelper = AuthHelper;
//# sourceMappingURL=AuthHelper.js.map