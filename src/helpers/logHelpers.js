const ProjectLog = require("../models/projectLog");

const updateProjectLog = async (id, title, action, date) => {
  const projectLog = new ProjectLog({
    id,
    log: [{ title, action, date }],
  });

  try {
    await projectLog.save();
    return { success: true, message: "Project log updated successfully" };
  } catch (error) {
    return {
      success: false,
      message: "Error updating project log",
      error: error.message,
    };
  }
};

module.exports = { updateProjectLog };
