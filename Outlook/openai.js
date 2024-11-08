require('dotenv').config();
const { OpenAI } = require('openai');

const openai = new OpenAI({
    apiKey: process.env.OPENAI_API_KEY,
});

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

let quotaExceeded = false;

const labelEmail = async (emailContent, retries = 3, delayTime = 10000) => {
    if (quotaExceeded) {
        console.warn("Quota exceeded, skipping API call.");
        return { label: "Quota exceeded", response: "Please try again later." };
    }

    try {
        const completion = await openai.chat.completions.create({
            model: "gpt-4o-mini",
            messages: [
                {
                    role: "system",
                    content: "You are an assistant that labels emails as 'interested', 'not interested', 'more info', or 'default'. Provide the label, followed by a response. End the mail with 'Best Regards, Aarush'. Add 'This mail is system generated, if you have any more specifications add them, we will revert back to you.'",
                },
                {
                    role: "user",
                    content: `Label the following email and generate a response:\n\n${emailContent} `,
                },
            ],
            max_tokens: 250,
        });

        const result = completion.choices[0].message.content.trim();
        const parts = result.split("\n").map(part => part.trim());
        const [label, ...responseParts] = parts;
        const response = responseParts.join("\n");

        console.log(`Email labeled as: ${label}`); // Log the label
        return { label, response };
    } catch (error) {
        if (error.response && error.response.status === 429) {
            console.warn("Rate limit exceeded, retrying...");
            quotaExceeded = true;
            await delay(delayTime);
            return await labelEmail(emailContent, retries - 1, delayTime * 2);
        } else {
            console.error("Error with OpenAI API:", error);
            return { label: "default", response: "Error processing request." };
        }
    }
};

module.exports = labelEmail;
