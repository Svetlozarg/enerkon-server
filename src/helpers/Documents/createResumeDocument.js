const path = require("path");
const { xmlToJs, writeToExcel } = require("../fileHelpers");
const { updateProjectLog } = require("../logHelpers");
const mongoose = require("mongoose");
const Document = require("../../models/documentModel");
const { uploadFileToDrive } = require("../FileStorage/fileStorageHelpers");

exports.createResumeDocument = async (projectId, filename, res) => {
  xmlToJs(filename, async (result) => {
    if (result.error) {
      res
        .status(result.error === "Error reading xml file" ? 500 : 400)
        .json(result);
    } else {
      const projectName = result.CalculationInput.General[0].ProjectName[0];

      // Contacts
      const sheetContacts = "Contacts";

      const heatedAreaCell = "C30";
      const heatedArea =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Area[0].HeatedArea[0];

      const heatedVolumeCell = "C31";
      const heatedVolume =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Area[0].HeatedVolume[0];

      const heatedAreaTwoCell = "C34";

      const heatedVolumeTwoCell = "C35";

      // Consumption
      const sheetConsumption = "Consumption";

      const cellC26 = "C26";
      const cellC26Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0].ResultNeededEnergyActual[0];

      const cellC28 = "C28";
      const cellC28Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HotWaterCalculations[0]
          .ResultNeededEnergyActual[0];

      const cellC30 = "C30";
      const cellC30Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].Lights[0].General[0]
          .DevicesNeededEnergyActual[0];

      const cellC31 = "C31";
      const cellC31DataBalanced =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].BalancedDevices[0]
          .General[0].DevicesNeededEnergyActual[0];
      const cellC31DataNonBalanced =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].NonBalancedDevices[0]
          .General[0].DevicesNeededEnergyActual[0];
      const cellC31Final = +cellC31DataBalanced + +cellC31DataNonBalanced;

      const cellE26 = "E26";
      const cellE26Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0]
          .ResultNeededEnergyBaseLine[0];

      const cellE28 = "E28";
      const cellE28Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HotWaterCalculations[0]
          .ResultNeededEnergyBaseLine[0];

      const cellE30 = "E30";
      const cellE30Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].Lights[0].General[0]
          .DevicesNeededEnergyActual[0];

      const cellE31 = "E31";
      const cellE31DataBalanced =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].BalancedDevices[0]
          .General[0].DevicesNeededEnergyActual[0];
      const cellE31DataNonBalanced =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].NonBalancedDevices[0]
          .General[0].DevicesNeededEnergyActual[0];
      const cellE31Final = +cellE31DataBalanced + +cellE31DataNonBalanced;

      const cellG11 = "G11";
      const cellG11Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0].UouterWallsSavings[0];

      const cellG26 = "G26";
      const cellG26Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0].ResultNeededEnergyESM[0];

      const cellG28 = "G28";
      const cellG28Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HotWaterCalculations[0]
          .ResultNeededEnergyESM[0];

      const cellG30 = "G30";
      const cellG30Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].Lights[0].General[0]
          .DevicesNeededEnergyActual[0];

      const cellG31 = "G31";
      const cellG31DataBalanced =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].BalancedDevices[0]
          .General[0].DevicesNeededEnergyActual[0];
      const cellG31DataNonBalanced =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].NonBalancedDevices[0]
          .General[0].DevicesNeededEnergyActual[0];
      const cellG31Final = +cellG31DataBalanced + +cellG31DataNonBalanced;

      const cellG35 = "G35";
      const cellG35Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0].UnontransparentSavings[0];

      const cellG47 = "G47";
      const cellG47Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0].UfloorSavings[0];

      const cellG59 = "G59";
      const cellG59Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0].UwindowsSavings[0];

      const cellG75 = "G75";
      const cellG75DataInfiltracion =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0].InfiltracionSavings[0];
      const cellG75DataSupplyNetEfficiency =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0]
          .SupplyNetEfficiencySavings[0];
      const cellG75DataAutomatic =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0].AutomaticSavings[0];
      const cellG75DataEnergyManagement =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0].EnergyManagementSavings[0];
      const cellG75DataGeneratorHeatEfficiency1 =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0]
          .GeneratorHeatEfficiency1Savings[0];
      const cellG75Final =
        +cellG75DataInfiltracion +
        +cellG75DataSupplyNetEfficiency +
        +cellG75DataAutomatic +
        +cellG75DataEnergyManagement +
        +cellG75DataGeneratorHeatEfficiency1;

      // Building Description
      const sheetBuildingDescription = "Building Description";

      const cellB12 = "B12";
      const cellB12Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Area[0].ZoneAreaElements[0].TotalRoofElements[0].Actual[0];

      const cellB13 = "B13";
      const cellB13Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Area[0].ZoneAreaElements[0].TotalFloorElements[0].Actual[0];

      const cellB16 = "B16";
      const cellB16Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0].InfiltracionActual[0];

      const cellB17 = "B17";
      const cellB17Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0].ProjectTemperatureActual[0];

      const cellB18 = "B18";
      const cellB18Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0]
          .NonProjectTemperatureActual[0];

      const cellB8 = "B8";
      const cellB8Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Area[0].ZoneAreaElements[0].TotalOuterWalls[0].Actual[0];

      const cellB9 = "B9";
      const cellB9Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Area[0].ZoneAreaElements[0].TotalWindows[0].Actual[0];

      const cellD100 = "D100";
      const cellD100Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HotWaterCalculations[0]
          .GeneratorHeatEfficiency1BaseLine[0];

      const cellD105 = "D105";
      const cellD105Data =
        result.CalculationInput.General[0].BuildingSavings[0].ZoneSaving[0]
          .Value[0];

      const cellD12 = "D12";
      const cellD12Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Roof[0].Current[0].NonTransparentU1[0];

      const cellD121 = "D121";
      const cellD121Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].Lights[0].General[0]
          .WorkScheduleBaseLine[0];

      const cellD122 = "D122";
      const cellD122Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].Lights[0].General[0]
          .DevicesNeededEnergyBaseLine[0];

      const cellD123 = "D123";
      const cellD123DataBalanced =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].BalancedDevices[0]
          .General[0].DevicesNeededEnergyBaseLine[0];
      const cellD123DataNonBalanced =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Area[0].HeatedArea[0];
      const cellD123Final = +cellD123DataBalanced + +cellD123DataNonBalanced;

      const cellD128 = "D128";
      const cellD128Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].BalancedDevices[0]
          .General[0].WorkScheduleBaseLine[0];

      const cellD129 = "D129";
      const cellD129Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].BalancedDevices[0]
          .General[0].DevicesNeededEnergyBaseLine[0];

      const cellD13 = "D13";
      const cellD13Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Floor[0].Current[0].FloorU1[0];

      const cellD130 = "D130";
      const cellD130DataBalanced =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].BalancedDevices[0]
          .General[0].DevicesNeededEnergyBaseLine[0];
      const cellD130DataNonBalanced =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Area[0].HeatedArea[0];
      const cellD130Final = +cellD130DataBalanced + +cellD130DataNonBalanced;

      const cellD134 = "D134";
      const cellD134Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].NonBalancedDevices[0]
          .General[0].WorkScheduleBaseLine[0];

      const cellD135 = "D135";
      const cellD135Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].NonBalancedDevices[0]
          .General[0].DevicesNeededEnergyBaseLine[0];

      const cellD136 = "D136";
      const cellD136DataBalanced =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].NonBalancedDevices[0]
          .General[0].DevicesNeededEnergyBaseLine[0];
      const cellD136DataNonBalanced =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Area[0].HeatedArea[0];
      const cellD136Final = +cellD136DataBalanced + +cellD136DataNonBalanced;

      const cellD16 = "D16";
      const cellD16Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0].InfiltracionBaseLine[0];

      const cellD17 = "D17";
      const cellD17Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0]
          .ProjectTemperatureBaseLine[0];

      const cellD18 = "D18";
      const cellD18Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0]
          .NonProjectTemperatureBaseLine[0];

      const cellD26 = "D26";
      const cellD26Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0]
          .GeneratorHeatEfficiency1BaseLine[0];

      const cellD29 = "D29";
      const cellD29Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0]
          .GeneratorHeatEfficiency1Actual[0];

      const cellD8 = "D8";
      const cellD8Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .NorthWalls[0].Current[0].OuterU1[0];

      const cellD9 = "D9";
      const cellD9Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .EastWalls[0].Current[0].AccumulateWindowU[0];

      const cellD98 = "D98";
      const cellD98Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HotWaterCalculations[0]
          .ResultSourceEnergyBaseLine[0];

      const cellE100 = "E100";
      const cellE100Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HotWaterCalculations[0]
          .GeneratorHeatEfficiency1ESM[0];

      const cellE105 = "E105";
      const cellE105Data =
        result.CalculationInput.General[0].BuildingSavings[0].ZoneSaving[0]
          .Value[0];

      const cellE121 = "E121";
      const cellE121Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].Lights[0].General[0]
          .WorkScheduleESM[0];

      const cellE122 = "E122";
      const cellE122Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].Lights[0].General[0]
          .DevicesNeededEnergyESM[0];

      const cellE123 = "E123";
      const cellE123DataBalanced =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].BalancedDevices[0]
          .General[0].DevicesNeededEnergyESM[0];
      const cellE123DataNonBalanced =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Area[0].HeatedArea[0];
      const cellE123Final = +cellE123DataBalanced + +cellE123DataNonBalanced;

      const cellE128 = "E128";
      const cellE128Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].BalancedDevices[0]
          .General[0].WorkScheduleESM[0];

      const cellE129 = "E129";
      const cellE129Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].BalancedDevices[0]
          .General[0].DevicesNeededEnergyESM[0];

      const cellE130 = "E130";
      const cellE130DataBalanced =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].BalancedDevices[0]
          .General[0].DevicesNeededEnergyESM[0];
      const cellE130DataNonBalanced =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Area[0].HeatedArea[0];
      const cellE130Final = +cellE130DataBalanced + +cellE130DataNonBalanced;

      const cellE134 = "E134";
      const cellE134Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].NonBalancedDevices[0]
          .General[0].WorkScheduleESM[0];

      const cellE135 = "E135";
      const cellE135Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].NonBalancedDevices[0]
          .General[0].DevicesNeededEnergyESM[0];

      const cellE136 = "E136";
      const cellE136DataBalanced =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].LightAndDevices[0].NonBalancedDevices[0]
          .General[0].DevicesNeededEnergyESM[0];
      const cellE136DataNonBalanced =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Area[0].HeatedArea[0];
      const cellE136Final = +cellE136DataBalanced + +cellE136DataNonBalanced;

      const cellE262 = "E26";
      const cellE26Data2 =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0]
          .GeneratorHeatEfficiency1ESM[0];

      const cellE29 = "E29";
      const cellE29Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0]
          .GeneratorHeatEfficiency1ESM[0];

      const cellE98 = "E98";
      const cellE98Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HotWaterCalculations[0]
          .ResultSourceEnergyESM[0];

      // Savings 2
      const sheetSavings2 = "Savings 2";

      const cellQ73 = "Q73";
      const cellQ73Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0].InfiltracionSavings[0];

      const cellQ75 = "Q75";
      const cellQ75Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0]
          .SupplyNetEfficiency2Savings[0];

      const cellQ77 = "Q77";
      const cellQ77Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0].Automatic2Savings[0];

      const cellQ79 = "Q79";
      const cellQ79Data =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0].EnergyManagement2Savings[0];

      const area =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].Heating[0]
          .Area[0].HeatedArea[0];
      const data1 =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].ZoneSavings[0]
          .ZoneSaving[0].ActualSaving[0];

      const baseData = +area * +data1.replace(",", ".");

      const cellG7Savings = "G7";
      const cellG7SavingsData = baseData;

      const cellG8Savings = "G8";
      const cellG8SavingsData = baseData;

      const cellG9Savings = "G9";
      const cellG9SavingsData = baseData;

      const cellG10Savings = "G10";
      const cellG10SavingsData = baseData;

      const cellG11Savings = "G11";
      const cellG11SavingsData = baseData;

      const cellG12Savings = "G12";
      const cellG12SavingsData = baseData;

      const cellG13Savings = "G13";
      const cellG13SavingsData = baseData;

      const cellG14Savings = "G14";
      const cellG14SavingsData = baseData;

      const cellG15Savings = "G15";
      const cellG15SavingsData = baseData;

      const cellG16Savings = "G16";
      const cellG16SavingsData = baseData;

      const cellG17Savings = "G17";
      const cellG17SavingsData = baseData;

      const data2 =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0].UnontransparentSavings[0];

      const baseDataTwo = +area * +data2.replace(",", ".");

      const cellG31Savings = "G31";
      const cellG31SavingsData = baseDataTwo;

      const cellG32Savings = "G32";
      const cellG32SavingsData = baseDataTwo;

      const cellG33Savings = "G33";
      const cellG33SavingsData = baseDataTwo;

      const cellG34Savings = "G34";
      const cellG34SavingsData = baseDataTwo;

      const cellG35Savings = "G35";
      const cellG35SavingsData = baseDataTwo;

      const cellG36Savings = "G36";
      const cellG36SavingsData = baseDataTwo;

      const cellG37Savings = "G37";
      const cellG37SavingsData = baseDataTwo;

      const cellG38Savings = "G38";
      const cellG38SavingsData = baseDataTwo;

      const cellG39Savings = "G39";
      const cellG39SavingsData = baseDataTwo;

      const cellG40Savings = "G40";
      const cellG40SavingsData = baseDataTwo;

      const cellG41Savings = "G41";
      const cellG41SavingsData = baseDataTwo;

      const data3 =
        result.CalculationInput.BuildingZones[0].BuildingZone[0].ZoneSavings[0]
          .ZoneSaving[0].ActualSaving[0];

      const baseDataThree = +area * +data3.replace(",", ".");

      const cellG43Savings = "G43";
      const cellG43SavingsData = baseDataThree;

      const cellG44Savings = "G44";
      const cellG44SavingsData = baseDataThree;

      const cellG45Savings = "G45";
      const cellG45SavingsData = baseDataThree;

      const cellG46Savings = "G46";
      const cellG46SavingsData = baseDataThree;

      const cellG47Savings = "G47";
      const cellG47SavingsData = baseDataThree;

      const cellG48Savings = "G48";
      const cellG48SavingsData = baseDataThree;

      const cellG49Savings = "G49";
      const cellG49SavingsData = baseDataThree;

      const cellG50Savings = "G50";
      const cellG50SavingsData = baseDataThree;

      const cellG51Savings = "G51";
      const cellG51SavingsData = baseDataThree;

      const cellG52Savings = "G52";
      const cellG52SavingsData = baseDataThree;

      const cellG53Savings = "G53";
      const cellG53SavingsData = baseDataThree;

      const data4 =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HeatingResult[0].UwindowsSavings[0];

      const baseDataFour = +area * +data4.replace(",", ".");

      const cellG55Savings = "G55";
      const cellG55SavingsData = baseDataFour;

      const cellG56Savings = "G56";
      const cellG56SavingsData = baseDataFour;

      const cellG57Savings = "G57";
      const cellG57SavingsData = baseDataFour;

      const cellG58Savings = "G58";
      const cellG58SavingsData = baseDataFour;

      const cellG59Savings = "G59";
      const cellG59SavingsData = baseDataFour;

      const cellG60Savings = "G60";
      const cellG60SavingsData = baseDataFour;

      const cellG61Savings = "G61";
      const cellG61SavingsData = baseDataFour;

      const cellG62Savings = "G62";
      const cellG62SavingsData = baseDataFour;

      const cellG63Savings = "G63";
      const cellG63SavingsData = baseDataFour;

      const cellG64Savings = "G64";
      const cellG64SavingsData = baseDataFour;

      const cellG65Savings = "G65";
      const cellG65SavingsData = baseDataFour;

      const data5 =
        result.CalculationInput.BuildingZones[0].BuildingZone[0]
          .HeatingCalculations[0].HotWaterCalculations[0]
          .ResultNeededEnergySavings[0];

      const baseDataFive = +area * +data5.replace(",", ".");

      const cellG83Savings = "G83";
      const cellG83SavingsData = baseDataFive;

      const cellG84Savings = "G84";
      const cellG84SavingsData = baseDataFive;

      const cellG85Savings = "G85";
      const cellG85SavingsData = baseDataFive;

      const cellG86Savings = "G86";
      const cellG86SavingsData = baseDataFive;

      const cellG87Savings = "G87";
      const cellG87SavingsData = baseDataFive;

      const cellG88Savings = "G88";
      const cellG88SavingsData = baseDataFive;

      const cellG89Savings = "G89";
      const cellG89SavingsData = baseDataFive;

      const cellG90Savings = "G90";
      const cellG90SavingsData = baseDataFive;

      const cellG91Savings = "G91";
      const cellG91SavingsData = baseDataFive;

      const cellG92Savings = "G92";
      const cellG92SavingsData = baseDataFive;

      const cellG93Savings = "G93";
      const cellG93SavingsData = baseDataFive;

      const xmlData = [
        // Contacts
        {
          sheetName: sheetContacts,
          cell: heatedAreaCell,
          value: heatedArea,
        },
        {
          sheetName: sheetContacts,
          cell: heatedVolumeCell,
          value: heatedVolume,
        },
        {
          sheetName: sheetContacts,
          cell: heatedAreaTwoCell,
          value: heatedArea,
        },
        {
          sheetName: sheetContacts,
          cell: heatedVolumeTwoCell,
          value: heatedVolume,
        },
        // Consumption
        {
          sheetName: sheetConsumption,
          cell: cellC26,
          value: cellC26Data,
        },
        {
          sheetName: sheetConsumption,
          cell: cellC28,
          value: cellC28Data,
        },
        {
          sheetName: sheetConsumption,
          cell: cellC30,
          value: cellC30Data,
        },
        {
          sheetName: sheetConsumption,
          cell: cellC31,
          value: cellC31Final,
        },
        {
          sheetName: sheetConsumption,
          cell: cellE26,
          value: cellE26Data,
        },
        {
          sheetName: sheetConsumption,
          cell: cellE28,
          value: cellE28Data,
        },
        {
          sheetName: sheetConsumption,
          cell: cellE30,
          value: cellE30Data,
        },
        {
          sheetName: sheetConsumption,
          cell: cellE31,
          value: cellE31Final,
        },
        {
          sheetName: sheetConsumption,
          cell: cellG11,
          value: cellG11Data,
        },
        {
          sheetName: sheetConsumption,
          cell: cellG26,
          value: cellG26Data,
        },
        {
          sheetName: sheetConsumption,
          cell: cellG28,
          value: cellG28Data,
        },
        {
          sheetName: sheetConsumption,
          cell: cellG30,
          value: cellG30Data,
        },
        {
          sheetName: sheetConsumption,
          cell: cellG31,
          value: cellG31Final,
        },
        {
          sheetName: sheetConsumption,
          cell: cellG35,
          value: cellG35Data,
        },
        {
          sheetName: sheetConsumption,
          cell: cellG47,
          value: cellG47Data,
        },
        {
          sheetName: sheetConsumption,
          cell: cellG59,
          value: cellG59Data,
        },
        {
          sheetName: sheetConsumption,
          cell: cellG75,
          value: cellG75Final,
        },
        // Building Description
        {
          sheetName: sheetBuildingDescription,
          cell: cellB12,
          value: cellB12Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellB13,
          value: cellB13Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellB16,
          value: cellB16Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellB17,
          value: cellB17Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellB18,
          value: cellB18Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellB8,
          value: cellB8Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellB9,
          value: cellB9Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD100,
          value: cellD100Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD105,
          value: cellD105Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD12,
          value: cellD12Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD121,
          value: cellD121Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD122,
          value: cellD122Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD123,
          value: cellD123Final,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD128,
          value: cellD128Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD129,
          value: cellD129Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD13,
          value: cellD13Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD130,
          value: cellD130Final,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD134,
          value: cellD134Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD135,
          value: cellD135Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD136,
          value: cellD136Final,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD16,
          value: cellD16Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD17,
          value: cellD17Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD18,
          value: cellD18Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD26,
          value: cellD26Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD29,
          value: cellD29Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD8,
          value: cellD8Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD9,
          value: cellD9Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellD98,
          value: cellD98Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellE100,
          value: cellE100Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellE105,
          value: cellE105Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellE121,
          value: cellE121Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellE122,
          value: cellE122Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellE123,
          value: cellE123Final,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellE128,
          value: cellE128Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellE129,
          value: cellE129Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellE130,
          value: cellE130Final,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellE134,
          value: cellE134Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellE135,
          value: cellE135Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellE136,
          value: cellE136Final,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellE262,
          value: cellE26Data2,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellE29,
          value: cellE29Data,
        },
        {
          sheetName: sheetBuildingDescription,
          cell: cellE98,
          value: cellE98Data,
        },
        // Savings 2
        {
          sheetName: sheetSavings2,
          cell: cellQ73,
          value: cellQ73Data,
        },
        {
          sheetName: sheetSavings2,
          cell: cellQ75,
          value: cellQ75Data,
        },
        {
          sheetName: sheetSavings2,
          cell: cellQ77,
          value: cellQ77Data,
        },
        {
          sheetName: sheetSavings2,
          cell: cellQ79,
          value: cellQ79Data,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG7Savings,
          value: cellG7SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG8Savings,
          value: cellG8SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG9Savings,
          value: cellG9SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG10Savings,
          value: cellG10SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG11Savings,
          value: cellG11SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG12Savings,
          value: cellG12SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG13Savings,
          value: cellG13SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG14Savings,
          value: cellG14SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG15Savings,
          value: cellG15SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG16Savings,
          value: cellG16SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG17Savings,
          value: cellG17SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG31Savings,
          value: cellG31SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG32Savings,
          value: cellG32SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG33Savings,
          value: cellG33SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG34Savings,
          value: cellG34SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG35Savings,
          value: cellG35SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG36Savings,
          value: cellG36SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG37Savings,
          value: cellG37SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG38Savings,
          value: cellG38SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG39Savings,
          value: cellG39SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG40Savings,
          value: cellG40SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG41Savings,
          value: cellG41SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG43Savings,
          value: cellG43SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG44Savings,
          value: cellG44SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG45Savings,
          value: cellG45SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG46Savings,
          value: cellG46SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG47Savings,
          value: cellG47SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG48Savings,
          value: cellG48SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG49Savings,
          value: cellG49SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG50Savings,
          value: cellG50SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG51Savings,
          value: cellG51SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG52Savings,
          value: cellG52SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG53Savings,
          value: cellG53SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG55Savings,
          value: cellG55SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG56Savings,
          value: cellG56SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG57Savings,
          value: cellG57SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG58Savings,
          value: cellG58SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG59Savings,
          value: cellG59SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG60Savings,
          value: cellG60SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG61Savings,
          value: cellG61SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG62Savings,
          value: cellG62SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG63Savings,
          value: cellG63SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG64Savings,
          value: cellG64SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG65Savings,
          value: cellG65SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG83Savings,
          value: cellG83SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG84Savings,
          value: cellG84SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG85Savings,
          value: cellG85SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG86Savings,
          value: cellG86SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG87Savings,
          value: cellG87SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG88Savings,
          value: cellG88SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG89Savings,
          value: cellG89SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG90Savings,
          value: cellG90SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG91Savings,
          value: cellG91SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG92Savings,
          value: cellG92SavingsData,
        },
        {
          sheetName: sheetSavings2,
          cell: cellG93Savings,
          value: cellG93SavingsData,
        },
      ];

      const currentDate = new Date();
      const day = currentDate.getDate().toString().padStart(2, "0");
      const month = (currentDate.getMonth() + 1).toString().padStart(2, "0");
      const year = currentDate.getFullYear().toString().slice(-2);

      const formattedDate = `${day}.${month}.${year}`;

      const resultFile = await writeToExcel(
        path.join("./uploads/", "resume.xlsx"),
        `Resume_Certificatе_${projectName}_${formattedDate}.xlsx`,
        xmlData
      );

      const newDocument = new Document({
        title: `Resume_Certificatе_${projectName}_${formattedDate}.xlsx`,
        document: {
          fileName: `Resume_Certificatе_${projectName}_${formattedDate}.xlsx`,
          originalName: `Resume_Certificatе_${projectName}_${formattedDate}.xlsx`,
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
        console.log(
          "\x1b[36m%s\x1b[0m",
          "Resume document successfully created"
        );
      } else {
        res.status(500).json({ error: resultFile.error });
      }
    }
  });
};
