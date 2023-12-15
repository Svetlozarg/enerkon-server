const fs = require("fs");
const xml2js = require("xml2js");
const ExcelJS = require("exceljs");
const path = require("path");

exports.getFileName = (filePath) => {
  return path.basename(filePath);
};

exports.getFileMimeType = (fileExtension) => {
  switch (fileExtension) {
    case "pdf":
      return "application/pdf";
    case "xml":
      return "application/xml";
    case "xlsx":
      return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    case "docs":
      return "application/vnd.google-apps.document";
    default:
      return "application/octet-stream";
  }
};

exports.xmlToJs = (filename, callback) => {
  fs.readFile("./uploads/" + filename, "utf-8", function (err, data) {
    if (err) {
      console.error("Error reading xml file:", err);
      callback({ error: "Error reading xml file" });
      return;
    }

    xml2js.parseString(data, (parseErr, result) => {
      if (parseErr) {
        console.error("Error parsing XML:", parseErr);
        callback({ error: "Invalid XML data" });
        return;
      }

      // Send the JSON response
      callback(result);
    });
  });
};

exports.writeToExcel = async (inputFilePath, filename, updates) => {
  try {
    // Load the existing Excel file
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(inputFilePath);

    for (const update of updates) {
      const { sheetName, cell, value } = update;

      // Get the worksheet by name
      const sheet = workbook.getWorksheet(sheetName);

      if (!sheet) {
        console.error(`Worksheet "${sheetName}" not found.`);
        continue;
      }

      // Update the cell value
      const targetCell = sheet.getCell(cell);
      targetCell.value = value;
    }

    // Create a new file with a different name
    const outputFileName = filename; // Change this to your desired output filename
    const outputPath = path.join("./uploads", outputFileName);

    // Save the modified workbook to the new file
    await workbook.xlsx.writeFile(outputPath);

    return {
      success: true,
      message: `Content added to Excel file "${outputFileName}" successfully`,
    };
  } catch (err) {
    console.error(err);
    return {
      success: false,
      error: "An error occurred while adding content to the Excel file",
    };
  }
};

exports.readFromExcel = async (filename) => {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(path.join("./uploads", filename));

    const worksheet = workbook.getWorksheet("Word");

    const data = [];

    worksheet.eachRow((row) => {
      const rowData = row.values.slice(1);
      data.push(rowData);
    });

    return data;
  } catch (error) {
    console.error("Error reading Excel file:", error.message);
    throw error;
  }
};
