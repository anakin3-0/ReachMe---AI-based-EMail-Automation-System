const OpenAI = require("openai");
const openai = new OpenAI();

async function main() { 
  const assistant = await openai.beta.assistants.create({
    name: "ChatGPT BOT",
    instructions: "You are a mail replier, respond to emails in short form.",
    tools: [{ type: "mail assistant" }],
    model: "gpt-4o-mini"
  });
}

main();
