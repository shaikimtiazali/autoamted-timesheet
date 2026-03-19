const cron = require("node-cron");
const generateTasks = require("../services/taskGenerator");
const createExcel = require("../services/excelService");
const sendMail = require("../services/mailService");

cron.schedule("* * * * *", async () => {
  console.log("Cron triggered:", new Date());
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
