const express = require("express");
const {
  getAllProjects,
  getProjectById,
  getProjectDocuments,
  deleteProject,
  getProjectLog,
  getProjectsAnalytics,
  createProject,
  updateProject,
} = require("../controllers/projectController");
const validateToken = require("../middleware/validateTokenHandler");
const upload = require("../middleware/upload");

const router = express.Router();

router.get("/projects", validateToken, getAllProjects);
router.get("/:id", validateToken, getProjectById);
router.get("/:id/documents", validateToken, getProjectDocuments);
router.post("/create", validateToken, upload.array("files", 2), createProject);
router.put("/update/:id", validateToken, updateProject);
router.delete("/delete", validateToken, deleteProject);
router.get("/log/:id", validateToken, getProjectLog);
router.get("/projects/analytics", validateToken, getProjectsAnalytics);

module.exports = router;
