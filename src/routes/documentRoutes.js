const express = require("express");
const {
  getAllDocuments,
  getDocumentById,
  createDocument,
  updateDocument,
  deleteDocument,
  downloadDocument,
  getPreviewLink,
} = require("../controllers/documentController");
const validateToken = require("../middleware/validateTokenHandler");
const upload = require("../middleware/upload");

const router = express.Router();

router.get("/documents", validateToken, getAllDocuments);
router.get("/:id", validateToken, getDocumentById);
router.post("/create", validateToken, upload.single("file"), createDocument);
router.put("/update/:id", validateToken, updateDocument);
router.delete("/delete", validateToken, deleteDocument);
router.get("/download/:fileName", validateToken, downloadDocument);
router.get("/preview/:fileName", validateToken, getPreviewLink);

module.exports = router;
