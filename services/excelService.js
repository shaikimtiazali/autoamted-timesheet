const Excel = require("exceljs");
const path = require("path");
require("dotenv").config();

//  Fixed schedule (NO randomness in time)
function getSchedule() {
  return [
    { start: "11:00 AM", end: "01:00 PM", type: "task" },
    { start: "03:00 PM", end: "05:00 PM", type: "task" },
    {
      start: "06:00 PM",
      end: "07:00 PM",
      task: "Daily Team Meeting with Amit Rai",
      type: "meeting",
    },
    {
      start: "07:30 PM",
      end: "08:00 PM",
      task: "Scrum call with Alan",
      type: "meeting",
    },
    { start: "08:00 PM", end: "09:00 PM", type: "task" },
  ];
}

//  Avoid duplicate tasks
function getUniqueTask(tasks, lastTask) {
  let newTask;
  do {
    newTask = tasks[Math.floor(Math.random() * tasks.length)].task;
  } while (newTask === lastTask);
  return newTask;
}

async function createExcel(tasks) {
  try {
    const templatePath = path.join(
      __dirname,
      "../templates/WFH Work Sheet.xlsx",
    );

    const today = new Date();
    const timestamp = new Date(Date.now()).toISOString().split("T")[0];

    const outputPath = path.join(
      __dirname,
      `../output/timesheet-${timestamp}.xlsx`,
    );

    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(templatePath);

    const sheet = workbook.getWorksheet(1);
    sheet.views = [
      {
        state: "normal",
        zoomScale: 100,
        showGridLines: true,
      },
    ];

    sheet.pageSetup = {
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: false,
      orientation: "landscape",
    };

    sheet.columns = [
      { key: "col1", width: 5 },
      { key: "date", width: 15 },
      { key: "start", width: 15 },
      { key: "end", width: 15 },
      { key: "col5", width: 10 },
      { key: "col6", width: 10 },
      { key: "task", width: 50 },
    ];

    //  Name
    sheet.getCell("C3").value = process.env.EMPLOYEE_NAME || "";
    // Signature
    sheet.getCell("C19").value = process.env.EMPLOYEE_NAME || "";

    //  Month
    const month = today.toLocaleString("default", { month: "long" });
    sheet.getCell("C5").value = month;

    const schedule = getSchedule();

    let rowStart = 10;
    let lastTask = "";

    schedule.forEach((entry, index) => {
      const row = sheet.getRow(rowStart + index);

      row.getCell(2).value = today;
      row.getCell(3).value = entry.start;
      row.getCell(4).value = entry.end;

      if (entry.type === "meeting") {
        row.getCell(7).value = entry.task;
      } else {
        const task = getUniqueTask(tasks, lastTask);
        lastTask = task;
        row.getCell(7).value = task;
      }

      //  Clean formatting
      row.getCell(2).numFmt = "dd-mmm-yyyy";
      row.getCell(3).alignment = { horizontal: "center" };
      row.getCell(4).alignment = { horizontal: "center" };
      row.getCell(7).alignment = { wrapText: true };

      row.commit();
    });

    await workbook.xlsx.writeFile(outputPath);

    console.log(" Excel generated (fixed schedule, 5 rows):", outputPath);

    return outputPath;
  } catch (error) {
    console.error(" Error in Excel generation:", error);
    throw error;
  }
}

module.exports = createExcel;
