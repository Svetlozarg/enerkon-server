const { google } = require("googleapis");
const fs = require("fs");
const path = require("path");
const { getFileMimeType, getFileName } = require("../fileHelpers");
const apikeys = require("../../../apikey.json");

const SCOPE = ["https://www.googleapis.com/auth/drive"];

const authorize = async () => {
  const jwtClient = new google.auth.JWT(
    apikeys.client_email,
    null,
    apikeys.private_key,
    SCOPE
  );

  await jwtClient.authorize();

  return jwtClient;
};

const uploadFile = (authClient, filePath) => {
  return new Promise((resolve, rejected) => {
    const drive = google.drive({ version: "v3", auth: authClient });
    const fileExtension = filePath.split(".").pop();
    const mimeType = getFileMimeType(fileExtension);

    const fileMetaData = {
      name: getFileName(filePath),
      parents: ["11WAtF3Ppi9Sfp2AkqDbMAldQPFgkigWy"],
    };

    drive.files.create(
      {
        resource: fileMetaData,
        media: {
          body: fs.createReadStream(filePath),
          mimeType: mimeType,
        },
        fields: "id",
      },
      function (error, file) {
        if (error) {
          return rejected(error);
        }
        resolve(file);
      }
    );
  });
};

exports.uploadFileToDrive = async (filePath) => {
  try {
    const authClient = await authorize();

    const file = await uploadFile(authClient, filePath);

    if (!file) {
      throw new Error("Error while uploading file");
    } else {
      console.log("%cFile uploaded successfully", "background-color: green");
    }
  } catch (error) {
    console.log(error);
  }
};

exports.downloadFileFromDrive = async (fileName, res) => {
  const authClient = await authorize();
  const drive = google.drive({ version: "v3", auth: authClient });

  const file = await drive.files.list({
    q: `name='${fileName}'`,
    fields: "files(id)",
  });

  if (file.data.files.length === 0) {
    throw new Error("File not found");
  }

  const fileId = file.data.files[0].id;

  const response = await drive.files.get(
    { fileId, alt: "media" },
    { responseType: "stream" }
  );

  const dest = fs.createWriteStream(`./uploads/${fileName}`);
  response.data
    .on("end", () => {
      console.log("Done downloading file.");
    })
    .on("error", (err) => {
      console.log("Error downloading file.", err);
    })
    .pipe(dest);

  const filePath = path.join("./uploads/", fileName);

  res.download(filePath);
};

exports.deleteFileFromDrive = async (fileName) => {
  const authClient = await authorize();
  const drive = google.drive({ version: "v3", auth: authClient });

  const file = await drive.files.list({
    q: `name='${fileName}'`,
    fields: "files(id)",
  });

  if (file.data.files.length === 0) {
    throw new Error("File not found");
  }

  const fileId = file.data.files[0].id;

  try {
    await drive.files.delete({ fileId });
    console.log("File deleted successfully.");
  } catch (error) {
    console.log("Error deleting file:", error);
  }
};

exports.getDocumentPreviewLink = async (fileName) => {
  const authClient = await authorize();
  const drive = google.drive({ version: "v3", auth: authClient });

  const file = await drive.files.list({
    q: `name='${fileName}'`,
    fields: "files(webViewLink)",
  });

  if (file.data.files.length === 0) {
    throw new Error("File not found");
  }

  const fileId = file.data.files[0].webViewLink;

  return fileId;
};
