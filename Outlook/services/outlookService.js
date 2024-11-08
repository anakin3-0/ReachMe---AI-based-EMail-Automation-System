const fetch = require('node-fetch');

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
        const errorData = await response.json(); // Get error details
        throw new Error(`Error fetching unread emails: ${response.statusText} - ${errorData.error.message}`);
    }

    const data = await response.json();
    console.log('Fetched unread emails:', data.value); // Log fetched emails
    return data.value; // This will be an array of email objects
};

// Function to send an email response
const sendEmailResponse = async (accessToken, emailId, responseMessage, recipientEmail) => {
    const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}/reply`, {
        method: 'POST',
        headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            message: {
                body: {
                    contentType: 'Text',
                    content: responseMessage,
                },
                toRecipients: [
                    {
                        emailAddress: {
                            address: recipientEmail,
                        },
                    },
                ],
            },
        }),
    });

    if (!response.ok) {
        const errorData = await response.json(); // Get error details
        throw new Error(`Error sending email response: ${response.statusText} - ${errorData.error.message}`);
    }

    console.log(`Sent response to ${recipientEmail}: ${responseMessage}`);
};

module.exports = { fetchUnreadEmails, sendEmailResponse };
