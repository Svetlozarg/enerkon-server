const path = require("path");
const Document = require("../../models/documentModel");
const mongoose = require("mongoose");
const { updateProjectLog } = require("../../helpers/logHelpers");
const { xmlToJs, writeToExcel } = require("../fileHelpers");
const { uploadFileToDrive } = require("../FileStorage/fileStorageHelpers");

exports.createKCCDocument = async (projectId, filename, res) => {
  xmlToJs(filename, async (result) => {
    if (result.error) {
      res
        .status(result.error === "Error reading xml file" ? 500 : 400)
        .json(result);
    } else {
      const sheetName = "Лист1";
      const nameCell = "A1:E1";
      const projectName = result.CalculationInput.General[0].ProjectName[0];

      const wallsCell = "D3";
      const totalWalls =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Area[0].ZoneAreaElements[0].TotalOuterWalls[0].Actual[0];

      const roofCell = "D4";
      const totalRoof =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Area[0].ZoneAreaElements[0].TotalRoofElements[0].Actual[0];

      const floorCell = "D5";
      const totalFloor =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Area[0].ZoneAreaElements[0].TotalFloorElements[0].Actual[0];

      const windowsCell = "D6";
      const totalWindows =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Area[0].ZoneAreaElements[0].TotalWindows[0].Actual[0];

      const xmlData = [
        {
          sheetName: sheetName,
          cell: nameCell,
          value: "КСС на сграда " + projectName,
        },
        { sheetName: sheetName, cell: wallsCell, value: totalWalls },
        { sheetName: sheetName, cell: roofCell, value: totalRoof },
        { sheetName: sheetName, cell: floorCell, value: totalFloor },
        { sheetName: sheetName, cell: windowsCell, value: totalWindows },
      ];

      const currentDate = new Date();
      const day = currentDate.getDate().toString().padStart(2, "0");
      const month = (currentDate.getMonth() + 1).toString().padStart(2, "0");
      const year = currentDate.getFullYear().toString().slice(-2);

      const formattedDate = `${day}.${month}.${year}`;

      const resultFile = await writeToExcel(
        path.join("./uploads/", "kcc.xlsx"),
        `КСС_${projectName}_${formattedDate}.xlsx`,
        xmlData
      );

      const newDocument = new Document({
        title: `КСС_${projectName}_${formattedDate}.xlsx`,
        document: {
          fileName: `КСС_${projectName}_${formattedDate}.xlsx`,
          originalName: `КСС_${projectName}_${formattedDate}.xlsx`,
          size: 12600,
        },
        project: projectId,
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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

      if (resultFile.success) {
        console.log("\x1b[36m%s\x1b[0m", "KCC document created successfully");
      } else {
        res.status(500).json({ error: resultFile.error });
      }
    }
  });
};
