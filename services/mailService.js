const nodemailer = require("nodemailer");

const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: process.env.EMAIL,
    pass: process.env.PASSWORD,
  },
});

async function sendMail(filePath) {
  const today = new Date();
  await transporter.sendMail({
    from: process.env.EMAIL,
    to: process.env.MANAGER_EMAIL,
    subject: "Daily Timesheet",
    text: `Please find the attached timesheet for today (${today.toDateString()}).`,
    attachments: [
      {
        filename: "timesheet.xlsx",
        path: filePath,
      },
    ],
  });
}

module.exports = sendMail;
