// app.js

require('dotenv').config();
const express = require('express');
const session = require('express-session');
const passport = require('passport');
const OAuth2Strategy = require('passport-oauth2');
const { Client } = require('@microsoft/microsoft-graph-client');
const labelEmail = require('./openai');
require('isomorphic-fetch');

const app = express();
const port = process.env.PORT || 3001;

// Middleware for session handling
app.use(
  session({
    secret: 'your_secret_key', // Replace with your own secret key
    resave: false,
    saveUninitialized: true,
  })
);

// Initialize Passport.js
app.use(passport.initialize());
app.use(passport.session());

// Configure Passport with OAuth2 strategy
passport.use(
  new OAuth2Strategy(
    {
      authorizationURL: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
      tokenURL: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      clientID: process.env.CLIENT_ID,
      clientSecret: process.env.CLIENT_SECRET,
      callbackURL: process.env.REDIRECT_URI,
      scope: [
        'openid',
        'profile',
        'User.Read',
        'Mail.Read',
        'Mail.Send',
        'offline_access',
      ],
    },
    (accessToken, refreshToken, params, profile, done) => {
      // Store accessToken and refreshToken in user object
      const user = {
        accessToken,
        refreshToken,
        expires_in: params.expires_in, // Token expiration time in seconds
      };
      return done(null, user);
    }
  )
);

// Serialize and deserialize user
passport.serializeUser((user, done) => {
  done(null, user);
});
passport.deserializeUser((user, done) => {
  done(null, user);
});

// Route to start authentication
app.get('/auth', passport.authenticate('oauth2'));

// Callback route
app.get(
  '/auth/callback',
  passport.authenticate('oauth2', { failureRedirect: '/' }),
  (req, res) => {
    // Authentication successful
    req.session.accessToken = req.user.accessToken; // Save access token in session
    req.session.refreshToken = req.user.refreshToken; // Save refresh token
    req.session.expiresAt = Date.now() + req.user.expires_in * 1000; // Calculate token expiration time
    res.redirect('/process-emails'); // Redirect to email processing
  }
);

// Middleware to ensure authentication
function ensureAuthenticated(req, res, next) {
  if (req.session.accessToken && Date.now() < req.session.expiresAt) {
    return next();
  } else if (req.session.refreshToken) {
    // Refresh the access token
    refreshAccessToken(req, res, next);
  } else {
    res.redirect('/auth');
  }
}

// Function to refresh access token
function refreshAccessToken(req, res, next) {
  const params = new URLSearchParams();
  params.append('client_id', process.env.CLIENT_ID);
  params.append('client_secret', process.env.CLIENT_SECRET);
  params.append('grant_type', 'refresh_token');
  params.append('refresh_token', req.session.refreshToken);
  params.append('scope', 'openid profile User.Read Mail.Read Mail.Send offline_access');

  fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
    method: 'POST',
    body: params,
  })
    .then((response) => response.json())
    .then((token) => {
      if (token.error) {
        console.error('Error refreshing access token:', token.error_description);
        res.redirect('/auth');
      } else {
        req.session.accessToken = token.access_token;
        req.session.refreshToken = token.refresh_token || req.session.refreshToken;
        req.session.expiresAt = Date.now() + token.expires_in * 1000;
        next();
      }
    })
    .catch((error) => {
      console.error('Error refreshing access token:', error);
      res.redirect('/auth');
    });
}

// Endpoint to fetch unread emails and process them
app.get('/process-emails', ensureAuthenticated, async (req, res) => {
  const accessToken = req.session.accessToken;

  try {
    const client = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });

    // Fetch unread emails
    const response = await client
      .api('/me/mailFolders/inbox/messages')
      .filter('isRead eq false')
      .select('id,subject,bodyPreview,body,from')
      .get();

    const emails = response.value;

    // Process each email
    for (const email of emails) {
      const emailContent = email.body.content;
      const senderEmail = email.from.emailAddress.address;
      const subject = email.subject;

      // Label the email and generate a response
      const { label, response: generatedResponse } = await labelEmail(emailContent);

      console.log(`Email from ${senderEmail} labeled as: ${label}`);

      // Send response back to sender
      await client.api('/me/sendMail').post({
        message: {
          subject: `Re: ${subject}`,
          body: {
            contentType: 'Text',
            content: generatedResponse,
          },
          toRecipients: [
            {
              emailAddress: {
                address: senderEmail,
              },
            },
          ],
        },
        saveToSentItems: 'true',
      });

      // Mark the email as read
      await client.api(`/me/messages/${email.id}`).patch({ isRead: true });
    }

    res.send('Processed unread emails and sent responses.');
  } catch (error) {
    console.error('processing emails:', error);
    res.status(500).send('processing emails.');
  }
});

// Start the Express app
app.listen(port, () => {
  console.log(`App listening on port ${port}`);
});
