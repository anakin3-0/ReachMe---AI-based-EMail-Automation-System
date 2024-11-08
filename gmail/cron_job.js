const { Queue, Worker, QueueScheduler } = require("bullmq");
const Redis = require("ioredis");
const { getEmails, addLabels, authorize } = require("./email");
const { labelEmail } = require("./openai");
const { sendEmail } = require("./sendEmail");

// Initialize Redis connection and queues
const connection = new Redis();
const emailQueue = new Queue("emailQueue", { connection });
const feedbackQueue = new Queue("feedbackQueue", { connection });
const replyQueue = new Queue("replyQueue", { connection });

// Initialize QueueSchedulers
new QueueScheduler("emailQueue", { connection });
new QueueScheduler("feedbackQueue", { connection });
new QueueScheduler("replyQueue", { connection });

// Worker to fetch emails and queue them for processing
new Worker(
  "emailQueue",
  async () => {
    const auth = await authorize();
    const emails = await getEmails(auth);

    for (const email of emails) {
      const feedbackJob = await feedbackQueue.add("processFeedback", { email });
      await replyQueue.add(
        "replyAndLabel",
        { email, feedbackJobId: feedbackJob.id },
        { parent: feedbackJob }
      );
    }
  },
  { connection }
);

// Worker to process feedback (label email and generate response)
new Worker(
  "feedbackQueue",
  async (job) => {
    const { email } = job.data;
    const response = await labelEmail(email.snippet); // Generate response using OpenAI
    const [label, reply] = response.split("\n", 2); // Split label and reply
    return { email, label, reply }; // Return the processed data for the next job
  },
  { connection }
);

// Worker to reply and label emails
new Worker(
  "replyQueue",
  async (job) => {
    const { email, feedbackJobId } = job.data;
    const feedbackJob = await feedbackQueue.getJob(feedbackJobId);
    const { label, reply } = feedbackJob.returnvalue;

    const auth = await authorize();
    await sendEmail(auth, email, reply); // Send the personalized reply
    await addLabels(auth, email.id, label); // Label the email in Gmail
  },
  { connection }
);

// Schedule the email fetching job every 15 minutes
emailQueue.add(
  "fetchEmails",
  {},
  {
    repeat: { every: 15 * 60 * 1000 }, // 15 minutes interval
  }
);
