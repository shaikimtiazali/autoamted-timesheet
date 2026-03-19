function generateTasks() {
  const tasks = require("../templates/tasks.json");
  return tasks
    .sort(() => 0.5 - Math.random())
    .slice(0, 3)
    .map((task) => ({ task }));
}

module.exports = generateTasks;
