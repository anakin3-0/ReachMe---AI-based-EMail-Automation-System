require("dotenv").config();
const { Client } = require("@microsoft/microsoft-graph-client");
const { InteractiveBrowserCredential } = require("@azure/identity");
const labelEmail = require("./labelEmail");
const sendEmail = require("./sendEmail");

const authenticate = async () => {
  const credential = new InteractiveBrowserCredential({
    tenantId: process.env.TENANT_ID,
  });

  const client = Client.initWithMiddleware({
    authProvider: (done) => {
      credential.getToken(process.env.MICROSOFT_GRAPH_SCOPE)
        .then(token => {
          done(null, token.token); // Pass the access token to the done callback
        })
        .catch(error => {
          console.error("Error acquiring token:", error); // Log the error for better debugging
          done(error, null); // Handle error case
        });
    },
  });

  return client;
};

const getEmails = async (client) => {
  const response = await client
    .api('/me/messages')
    .filter("isRead eq false")
    .top(10)
    .get();
  return response.value;
};

const processEmails = async () => {
  try {
    const client = await authenticate();
    const emails = await getEmails(client);

    for (const email of emails) {
      const emailContent = email.body.content;
      const { label, response } = await labelEmail(emailContent);
      await sendEmail(email.sender.emailAddress.address, response);
      
      console.log(`Processed email ID: ${email.id}, Label: ${label}`);
    }
  } catch (error) {
    console.error("Error processing emails:", error); // Handle any errors during email processing
  }
};

// processEmails().catch(console.error);


module.exports = processEmails;