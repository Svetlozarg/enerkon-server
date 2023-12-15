const fs = require("fs");
const path = require("path");
const Docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");
const { readFromExcel } = require("../fileHelpers");
const mongoose = require("mongoose");
const { updateProjectLog } = require("../logHelpers");
const Document = require("../../models/documentModel");
const { uploadFileToDrive } = require("../FileStorage/fileStorageHelpers");

// Function to create a Word Report Document
exports.createReportDocument = async (
  masteFileName,
  projectName,
  projectId
) => {
  try {
    // Get masterfile data
    const masterFileData = await readFromExcel(masteFileName);

    // Read the template content
    const templatePath = path.resolve("./uploads", "report.docx");
    const content = fs.readFileSync(templatePath, "binary");

    // Create a PizZip instance
    const zip = new PizZip(content);

    // Create a Docxtemplater instance
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    // Build render data object
    const renderData = {};
    masterFileData.forEach((item) => {
      const [key, value] = item;
      renderData[key.toLowerCase()] = value;
    });

    // Render the template with masterFileData
    doc.render(renderData);

    // Generate the output buffer
    const buf = doc.getZip().generate({
      type: "nodebuffer",
      compression: "DEFLATE",
    });

    const currentDate = new Date();
    const day = currentDate.getDate().toString().padStart(2, "0");
    const month = (currentDate.getMonth() + 1).toString().padStart(2, "0");
    const year = currentDate.getFullYear().toString().slice(-2);

    const formattedDate = `${day}.${month}.${year}`;

    // Write the output to a file
    const outputPath = path.resolve(
      "./uploads",
      `Доклад_${projectName}_${formattedDate}.doc`
    );
    fs.writeFileSync(outputPath, buf);

    const newDocument = new Document({
      title: `Доклад_${projectName}_${formattedDate}.doc`,
      document: {
        fileName: `Доклад_${projectName}_${formattedDate}.doc`,
        originalName: `Доклад_${projectName}_${formattedDate}.doc`,
        size: 12600,
      },
      project: projectId,
      type: "application/msword",
    });

    const createdDocument = await newDocument.save();

    updateProjectLog(
      new mongoose.Types.ObjectId(projectId),
      createdDocument.title,
      "Файлът е създаден",
      createdDocument.updatedAt
    );

    const filePath = path.join("./uploads/", createdDocument.title);

    await uploadFileToDrive(filePath);

    console.log("\x1b[36m%s\x1b[0m", "Report document created successfully!");
  } catch (error) {
    console.error("Error creating report document:", error);
  }
};
