require("dotenv").config();
const { Client } = require("@microsoft/microsoft-graph-client");
const { InteractiveBrowserCredential } = require("@azure/identity");

const authenticate = async () => {
  const credential = new InteractiveBrowserCredential({
    tenantId: process.env.TENANT_ID,
  });
  const client = Client.initWithMiddleware({
    authProvider: (done) => {
      credential.getToken(process.env.MICROSOFT_GRAPH_SCOPE)
        .then(token => done(null, token.token))
        .catch(error => done(error, null));
    },
  });

  return client;
};

const sendEmail = async (to, body) => {
  const client = await authenticate();
  
  await client.api('/me/sendMail').post({
    message: {
      subject: "Automated Response",
      body: {
        contentType: "Text",
        content: body,
      },
      toRecipients: [
        {
          emailAddress: {
            address: to,
          },
        },
      ],
    },
  });
};

module.exports = sendEmail;
