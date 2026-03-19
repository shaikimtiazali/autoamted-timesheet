const Excel = require("exceljs");
const path = require("path");
require("dotenv").config();

//  Format hour
function formatTime(hour, minutes = 0) {
  const suffix = hour >= 12 ? "PM" : "AM";
  let h = hour > 12 ? hour - 12 : hour;
  if (h === 0) h = 12;
  return `${h}:${minutes.toString().padStart(2, "0")} ${suffix}`;
}

//  Convert to comparable number
function toMinutes(hour, minutes = 0) {
  return hour * 60 + minutes;
}

//  Fixed meetings
function getFixedMeetings() {
  return [
    {
      startMin: toMinutes(18, 0),
      endMin: toMinutes(19, 0),
      start: "6:00 PM",
      end: "7:00 PM",
      task: "Daily Team Meeting with Amit Rai",
    },
    {
      startMin: toMinutes(19, 30),
      endMin: toMinutes(20, 0),
      start: "7:30 PM",
      end: "8:00 PM",
      task: "Scrum call with Alan",
    },
  ];
}

//  Generate working slots excluding lunch + meetings
function generateTimeSlots() {
  const slots = [];

  const blocks = [
    { start: 11, end: 14 }, // 11–2
    { start: 15, end: 18 }, // 3–6
    { start: 19, end: 20 }, // 7–8 (after meeting)
  ];

  blocks.forEach((block) => {
    let current = block.start;

    while (current < block.end) {
      const duration = Math.floor(Math.random() * 2) + 1;
      let end = current + duration;

      if (end > block.end) end = block.end;

      slots.push({
        start: formatTime(current),
        end: formatTime(end),
        startMin: toMinutes(current),
        endMin: toMinutes(end),
      });

      current = end;
    }
  });

  return slots;
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

    //  Generate slots + fixed meetings
    const randomSlots = generateTimeSlots();
    const meetings = getFixedMeetings();

    const allSlots = [
      ...randomSlots.map((s) => ({ ...s, type: "task" })),
      ...meetings.map((m) => ({ ...m, type: "meeting" })),
    ];

    // Sort by time
    allSlots.sort((a, b) => a.startMin - b.startMin);

    let rowStart = 10;
    let lastTask = "";

    allSlots.forEach((slot, index) => {
      const row = sheet.getRow(rowStart + index);

      row.getCell(2).value = today;
      row.getCell(3).value = slot.start;
      row.getCell(4).value = slot.end;

      if (slot.type === "meeting") {
        row.getCell(7).value = slot.task;
      } else {
        const task = getUniqueTask(tasks, lastTask);
        lastTask = task;
        row.getCell(7).value = task;
      }

      row.commit();
    });

    await workbook.xlsx.writeFile(outputPath);

    console.log("Excel generated:", outputPath);

    return outputPath;
  } catch (error) {
    console.error("Error in Excel generation:", error);
    throw error;
  }
}

module.exports = createExcel;
