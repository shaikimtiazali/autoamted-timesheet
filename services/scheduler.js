const cron = require("node-cron");
const generateTasks = require("../services/taskGenerator");
const createExcel = require("../services/excelService");
const sendMail = require("../services/mailService");

// "0 19 * * 1-5" Monday - Friday at 7 PM
cron.schedule("* * * * *", async () => {
  console.log("Running Timesheet Job");
  try {
    const tasks = generateTasks();
    const file = await createExcel(tasks);
    await sendMail(file);
    console.log("Timesheet sent successfully");
  } catch (error) {
    console.error("Error sending timesheet:", error);
  }
});
