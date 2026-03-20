// require("dotenv").config();

// require("./scheduler/scheduler");

// console.log("Timesheet automation started ...");

// Above configuration is commented out for testing purposes. Uncomment it when you want to run the scheduler and load environment variables from a .env file every time.

(async () => {
  try {
    const tasks = require("./services/taskGenerator")();
    const createExcel = require("./services/excelService");
    const sendMail = require("./services/mailService");

    const file = await createExcel(tasks);
    await sendMail(file);

    console.log("Timesheet sent successfully");
    process.exit(0);
  } catch (err) {
    console.error(err);
    process.exit(1);
  }
})();
