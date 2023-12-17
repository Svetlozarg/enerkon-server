const asyncHandler = require("express-async-handler");
const mongoose = require("mongoose");
const Document = require("../models/documentModel");
const {
  uploadFileToDrive,
  downloadFileFromDrive,
  deleteFileFromDrive,
  getDocumentPreviewLink,
} = require("../helpers/FileStorage/fileStorageHelpers");
const { updateProjectLog } = require("../helpers/logHelpers");
const path = require("path");
const fs = require("fs");

//@desc Get all documents
//?@route GET /api/document/documents
//@access private
exports.getAllDocuments = asyncHandler(async (req, res, next) => {
  const documents = await Document.find();

  res.status(200).json({
    success: true,
    data: documents,
  });
});

//@desc Get a document by ID
//?@route GET /api/document/:id
//@access private
exports.getDocumentById = asyncHandler(async (req, res, next) => {
  const document = await Document.findById(req.params.id);

  if (!document) {
    res.status(404);
    throw new Error("Document not found");
  }

  res.status(200).json({
    success: true,
    data: document,
  });
});

//@desc Create a document
//!@route POST /api/document/create
//@access private
exports.createDocument = asyncHandler(async (req, res, next) => {
  const { projectId, filename } = req.body;
  const file = req.file;

  if (!filename || !file) {
    res.status(400);
    throw new Error("No file name or file uploaded");
  }

  const newDocument = new Document({
    title: filename,
    document: {
      fileName: req.customName,
      originalName: filename,
      size: file.size,
    },
    project: projectId,
    type: file.mimetype,
  });

  const createdDocument = await newDocument.save();

  updateProjectLog(
    new mongoose.Types.ObjectId(projectId),
    createdDocument.title,
    "Файлът е създаден",
    createdDocument.updatedAt
  );

  const filePath = path.join("./uploads/", req.customName);

  await uploadFileToDrive(filePath);

  res.status(201).json({
    success: true,
    message: "Document created successfully",
    data: createdDocument,
  });
});

//@desc Update a document
//!@route PUT /api/document/update/:id
//@access private
exports.updateDocument = asyncHandler(async (req, res, next) => {
  const documentId = req.params.id;
  const updatedData = req.body;

  const updatedDocument = await Document.findByIdAndUpdate(
    documentId,
    updatedData,
    { new: true }
  );

  if (!updatedDocument) {
    res.status(404);
    throw new Error("Document not found");
  }

  let logMessage = "Document updated successfully";

  if (updatedData.title && updatedData.status) {
    logMessage = "Заглавието и статуса са променени";
  } else if (updatedData.title) {
    logMessage = "Заглавието на файла е променено";
  } else if (updatedData.status) {
    logMessage = "Статусът на файла е променен";
  }

  updateProjectLog(
    new mongoose.Types.ObjectId(updatedDocument.project),
    updatedDocument.title,
    logMessage,
    updatedDocument.updatedAt
  );

  res.json({
    success: true,
    message: logMessage,
    data: updatedDocument,
  });
});

//@desc Delete a document
//!@route DELETE /api/document/delete
//@access private
exports.deleteDocument = asyncHandler(async (req, res, next) => {
  const { id, fileName } = req.body;

  const deleteDocument = await Document.findByIdAndDelete(id);

  if (!deleteDocument) {
    res.status(404);
    res.json({
      success: false,
      message: "Document not found",
    });
    throw new Error("Document not found");
  }

  updateProjectLog(
    new mongoose.Types.ObjectId(deleteDocument.project),
    deleteDocument.title,
    "Файлът е изтрит",
    deleteDocument.updatedAt
  );

  deleteFileFromDrive(fileName);

  res.json({
    success: true,
    message: "Document deleted successfully",
  });
});

//@desc Download a document
//?@route GET /api/document/download/:fileName
//@access private
exports.downloadDocument = asyncHandler(async (req, res, next) => {
  const { fileName } = req.params;

  await downloadFileFromDrive(fileName, res);

  setTimeout(() => {
    const filePath = path.join("./uploads/", fileName);

    if (fs.existsSync(filePath)) {
      res.download(filePath, fileName);
    } else {
      res.status(404).json({
        success: false,
        message: "File not found",
      });
    }
  }, 1000);
});

//@desc Get document link to google drive
//?@route GET /api/document/preview/:fileName
//@access private
exports.getPreviewLink = asyncHandler(async (req, res, next) => {
  const { fileName } = req.params;

  const previewLink = await getDocumentPreviewLink(fileName);

  res.status(200).json({
    success: true,
    data: previewLink,
  });
});
