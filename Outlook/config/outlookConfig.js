require('dotenv').config();
const querystring = require('querystring');
const fetch = require('node-fetch'); // Make sure to install node-fetch if you haven't

const clientId = process.env.OUTLOOK_CLIENT_ID;
const clientSecret = process.env.OUTLOOK_CLIENT_SECRET;
const tenantId = process.env.OUTLOOK_TENANT_ID;
const redirectUri = process.env.OUTLOOK_REDIRECT_URI;

// Function to get the authorization URL
const getAuthorizationUrl = () => {
    const params = {
        client_id: clientId,
        response_type: 'code',
        redirect_uri: redirectUri,
        response_mode: 'query',
        scope: 'openid email profile offline_access User.Read Mail.ReadWrite'
    };
    const queryString = querystring.stringify(params);
    return `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?${queryString}`;
};

// Function to exchange auth code for access token
const getAccessToken = async (authCode) => {
    const response = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: querystring.stringify({
            client_id: clientId,
            client_secret: clientSecret,
            code: authCode,
            redirect_uri: redirectUri,
            grant_type: 'authorization_code',
        }),
    });

    if (!response.ok) {
        const errorData = await response.json(); // Get error details
        throw new Error(`Error getting access token: ${response.statusText} - ${errorData.error}`);
    }

    const data = await response.json();
    console.log('Access Token:', data.access_token); // Log the access token
    return data.access_token;
};

// Function to fetch unread emails
const fetchUnreadEmails = async (accessToken) => {
    const response = await fetch('https://graph.microsoft.com/v1.0/me/messages?$filter=isRead eq false', {
        method: 'GET',
        headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
        },
    });

    if (!response.ok) {
        const errorData = await response.json(); // Get error details from the response
        throw new Error(`Error fetching unread emails: ${response.statusText} - ${errorData.error.message}`);
    }

    const data = await response.json();
    console.log('Fetched unread emails:', data.value); // Log fetched emails
    return data.value; // This will be an array of email objects
};

module.exports = { getAuthorizationUrl, getAccessToken, fetchUnreadEmails };
