const express = require('express');
const axios = require('axios');
const session = require('express-session');
const OpenAI = require('openai'); // Use the old syntax for OpenAI API

const app = express();
const PORT = 3000;

const CLIENT_ID='***';// Replace with your client ID
const CLIENT_SECRET = '***'; // Replace with your client secret
const TENANT_ID = '***'; // Replace with your tenant ID

app.use(session({
  secret: 'your_secret_key',
  resave: false,
  saveUninitialized: true,
}));

// Auth route
app.get('/auth/outlook', (req, res) => {
  const redirectUri = `http://localhost:${PORT}/auth/outlook/callback`;
  const authUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?client_id=${CLIENT_ID}&response_type=code&redirect_uri=${encodeURIComponent(redirectUri)}&response_mode=query&scope=https://graph.microsoft.com/.default`;
  
  res.redirect(authUrl);
});

// Callback route
app.get('/auth/outlook/callback', (req, res) => {
  const code = req.query.code;

  if (!code) {
    return res.send('Authorization code not provided.');
  }

  // Ask user for the code in terminal
  console.log('Authorization Code:', code);
  // Process the code to get the access token
  getAccessToken(code)
    .then(token => {
      console.log('Access Token:', token);
      res.send('Authentication successful! Check your console for the access token.');
    })
    .catch(err => {
      console.error('Error processing the token:', err);
      res.send('Error processing the token. Check console for details.');
    });
});

// Function to get access token
async function getAccessToken(code) {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  
  const params = new URLSearchParams();
  params.append('grant_type', 'authorization_code');
  params.append('client_id', CLIENT_ID);
  params.append('client_secret', CLIENT_SECRET);
  params.append('code', code);
  params.append('redirect_uri', `http://localhost:${PORT}/auth/outlook/callback`);
  params.append('scope', 'https://graph.microsoft.com/.default');

  const response = await axios.post(tokenUrl, params);
  return response.data.access_token;
}

// Process emails route
app.get('/process-emails', async (req, res) => {
  const accessToken = req.session.accessToken; // Store access token in session after getting it

  if (!accessToken) {
    return res.status(401).send('Access token is required for processing emails.');
  }

  try {
    // Fetch emails using Microsoft Graph API
    const response = await axios.get('https://graph.microsoft.com/v1.0/me/messages', {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });
    res.json(response.data.value);
  } catch (error) {
    console.error('Error fetching emails:', error);
    res.status(500).send('Error fetching emails. Check console for details.');
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
