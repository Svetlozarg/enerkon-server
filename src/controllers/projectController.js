const asyncHandler = require("express-async-handler");
const Project = require("../models/projectModel");
const Document = require("../models/documentModel");
const mongoose = require("mongoose");
const { createKCCDocument } = require("../helpers/Documents/createKCCDocument");
const {
  createReportDocument,
} = require("../helpers/Documents/createReportDocument");
const {
  createResumeDocument,
} = require("../helpers/Documents/createResumeDocument");
const { updateProjectLog } = require("../helpers/logHelpers");
const {
  uploadFileToDrive,
  deleteFileFromDrive,
  downloadFileFromDrive,
} = require("../helpers/FileStorage/fileStorageHelpers");
const path = require("path");

//@desc Get all projects
//?@route GET /api/project/projects
//@access private
exports.getAllProjects = asyncHandler(async (req, res, next) => {
  const projects = await Project.find();
  res.status(200).json({ success: true, data: projects });
});

//@desc Get a project by ID
//?@route GET /api/project/:id
//@access private
exports.getProjectById = asyncHandler(async (req, res, next) => {
  const project = await Project.findById(req.params.id);
  if (!project) {
    res.status(404);
    throw new Error("Project not found");
  }
  res.status(200).json({ success: true, data: project });
});

//@desc Get project documents
//?@route GET /api/project/:id/documents
//@access private
exports.getProjectDocuments = asyncHandler(async (req, res, next) => {
  const project = await Project.findById(req.params.id);
  if (!project) {
    res.status(404);
    throw new Error("Project not found");
  }
  const documents = await Document.find({
    project: new mongoose.Types.ObjectId(req.params.id),
  });
  res.status(200).json({ success: true, data: documents });
});

//@desc Get a project log
//?@route GET /api/project/log/:id
//@access private
exports.getProjectLog = asyncHandler(async (req, res, next) => {
  const projectLogs = await ProjectLog.find({ id: req.params.id });
  if (!projectLogs) {
    res.status(404);
    throw new Error("Project log not found");
  }
  res.status(200).json({ success: true, data: projectLogs });
});

//@desc Get projects analytics
//?@route GET /api/project/analytics
//@access private
exports.getProjectsAnalytics = asyncHandler(async (req, res, next) => {
  const analytics = await Project.aggregate([
    {
      $group: {
        _id: {
          month: { $month: "$createdAt" },
          year: { $year: "$createdAt" },
        },
        paidCount: {
          $sum: { $cond: [{ $eq: ["$status", "Paid"] }, 1, 0] },
        },
        unpaidCount: {
          $sum: { $cond: [{ $eq: ["$status", "Unpaid"] }, 1, 0] },
        },
      },
    },
    {
      $sort: { "_id.year": 1, "_id.month": 1 },
    },
  ]);

  const paidArray = Array(12).fill(0);
  const unpaidArray = Array(12).fill(0);

  analytics.forEach((entry) => {
    const monthIndex = entry._id.month - 1;
    paidArray[monthIndex] += entry.paidCount;
    unpaidArray[monthIndex] += entry.unpaidCount;
  });

  res.json({
    success: true,
    data: {
      paid: paidArray,
      unpaid: unpaidArray,
    },
  });
});

//@desc Get projects analytics
//?@route GET /api/project/analytics
//@access private

//@desc Create a project
//!@route POST /api/project/create
//@access private
exports.createProject = asyncHandler(async (req, res, next) => {
  const { title } = req.body;
  const files = req.files;

  if (!title) {
    throw new Error("Project title is required");
  }

  if (files.length === 0) {
    throw new Error("No files uploaded");
  }

  const newProject = new Project({
    title,
    documents: files.map((file) => ({
      fileName: file.filename,
      originalName: file.originalname,
      type: file.mimetype,
      size: file.size,
    })),
  });

  const createdProject = await newProject.save();

  const documentPromises = files.map((file) => {
    const newDocument = new Document({
      title: file.originalname,
      document: {
        id: file.id,
        fileName: file.filename,
        originalName: file.originalname,
        size: file.size,
      },
      project: createdProject.id,
      type: file.mimetype,
      status: "Finished",
      default: true,
    });

    return newDocument.save();
  });

  files.map((file) => {
    const filePath = path.join("./uploads/", file.filename);

    uploadFileToDrive(filePath);
  });

  const createdDocuments = await Promise.all(documentPromises);

  const logPromises = createdDocuments.map((document) => {
    return updateProjectLog(
      new mongoose.Types.ObjectId(newProject._id),
      document.title,
      "Документът е създаден",
      document.createdAt
    );
  });

  await Promise.all(logPromises);

  createKCCDocument(
    createdDocuments[0].project,
    createdDocuments[0].document.fileName,
    res
  );

  createReportDocument(
    createdDocuments[1].document.fileName,
    title,
    createdProject.id
  );

  createResumeDocument(
    createdDocuments[0].project,
    createdDocuments[0].document.fileName,
    res
  );

  res.status(201).json({
    success: true,
    message: "Project created successfully",
    data: createdProject,
  });
});

//@desc Update a project
//!@route PUT /api/project/update/:id
//@access private
exports.updateProject = asyncHandler(async (req, res, next) => {
  const projectId = req.params.id;
  const updatedData = req.body;

  const updatedProject = await Project.findByIdAndUpdate(
    projectId,
    updatedData,
    { new: true }
  );

  if (!updatedProject) {
    res.status(404);
    throw new Error("Project not found");
  }

  let logMessage = "";

  if (updatedData.title && updatedData.status) {
    logMessage = "Заглавието и статуса са променени";
  } else if (updatedData.title) {
    logMessage = "Заглавието на проекта е променено";
  } else if (updatedData.status) {
    logMessage = "Статусът на проекта е променен";
  } else if (updatedData.favourite) {
    logMessage = "Проектът е добавен в любими";
  } else {
    logMessage = "Проектът е премахнат от любими";
  }

  updateProjectLog(
    new mongoose.Types.ObjectId(projectId),
    updatedProject.title,
    logMessage,
    updatedProject.updatedAt
  );

  res.json({
    success: true,
    message: "Project updated successfully",
    data: updatedProject,
  });
});

//@desc Delete a project
//!@route DELETE /api/project/delete
//@access private
exports.deleteProject = asyncHandler(async (req, res, next) => {
  const { id } = req.body;

  const project = await Project.findById(id);

  if (!project) {
    return res.status(404).json({ message: "Project not found" });
  }

  const documents = await Document.find({ project: id });

  if (documents) {
    const deleteProject = await Project.findByIdAndDelete(id);
    if (!deleteProject) {
      res.status(404);
      throw new Error("Project not found");
    }

    const projectLogs = await Project.find({ id: id });

    for (const log of projectLogs) {
      await Project.findByIdAndDelete(log._id);
    }

    for (const document of documents) {
      const deleteDocument = await Document.findByIdAndDelete(
        new mongoose.Types.ObjectId(document.id)
      );

      deleteFileFromDrive(document.document.fileName);

      if (!deleteDocument) {
        res.status(404);
        res.json({
          success: false,
          message: "Document not found",
        });
        throw new Error("Document not found");
      }
    }
  }

  res.json({
    success: true,
    message: "Project deleted successfully",
  });
});

//@desc Recreate a project
//?@route GET /api/project/recreate/:project
//@access private
exports.recreateProjectDocuments = asyncHandler(async (req, res, next) => {
  const projectId = req.params.project;

  const project = await Project.findById(projectId);

  if (!project) {
    res.status(404);
    throw new Error("Project not found");
  }

  const defaultDocuments = await Document.find({
    project: projectId,
    default: true,
  });

  if (defaultDocuments) {
    for (const document of defaultDocuments) {
      const downloadDocument = await Document.findById(
        new mongoose.Types.ObjectId(document.id)
      );

      if (!downloadDocument) {
        res.status(404);
        res.json({
          success: false,
          message: "Document not found",
        });
        throw new Error("Document not found");
      }

      await downloadFileFromDrive(document.document.fileName);
    }
  }

  const documents = await Document.find({ project: projectId, default: false });

  if (documents) {
    for (const document of documents) {
      if (document.default) {
        continue;
      }

      const deleteDocument = await Document.findByIdAndDelete(
        new mongoose.Types.ObjectId(document.id)
      );

      deleteFileFromDrive(document.document.fileName);

      if (!deleteDocument) {
        res.status(404);
        res.json({
          success: false,
          message: "Document not found",
        });
        throw new Error("Document not found");
      }
    }
  }

  createKCCDocument(
    defaultDocuments[0].project,
    defaultDocuments[0].document.fileName,
    res
  );

  createReportDocument(
    defaultDocuments[1].document.fileName,
    project.title,
    defaultDocuments[1].project
  );

  createResumeDocument(
    defaultDocuments[0].project,
    defaultDocuments[0].document.fileName,
    res
  );

  updateProjectLog(
    new mongoose.Types.ObjectId(projectId),
    project.title,
    "Проектът е пресъздаден",
    Date.now()
  );

  res.json({
    success: true,
    message: "Project recreated successfully",
  });
});
