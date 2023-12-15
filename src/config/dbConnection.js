const mongoose = require("mongoose");
const fs = require("fs");
const path = require("path");
const { Buffer } = require("buffer");
const fetch = require("node-fetch");

const FILE_URLS = [
  "https://ik.imagekit.io/obelussoft/Enerkon/kcc.xlsx",
  "https://ik.imagekit.io/obelussoft/Enerkon/report.docx",
  "https://ik.imagekit.io/obelussoft/Enerkon/resume.xlsx",
];

const connectDb = async () => {
  try {
    const connect = await mongoose.connect(process.env.CONNECTION_STRING, {
      useNewUrlParser: true,
      useUnifiedTopology: true,
    });

    const uploadsFolderPath = path.resolve("./uploads");

    if (!fs.existsSync(uploadsFolderPath)) {
      await fs.mkdir(uploadsFolderPath);
      console.log("Uploads folder created.");
    } else {
      console.log("Uploads folder already exists.");
    }

    // Download files and save to uploads folder using fetch
    await Promise.all(
      FILE_URLS.map(async (fileUrl) => {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();

        // Convert ArrayBuffer to Buffer using Buffer.from
        const buffer = Buffer.from(arrayBuffer);

        // Extract the filename from the URL or provide a custom filename
        const fileName = path.basename(fileUrl);

        // Save the file to the uploads folder
        const filePath = path.join(uploadsFolderPath, fileName);
        fs.writeFileSync(filePath, buffer, (err) => {
          if (err) throw err;
          console.log(`File downloaded and saved: ${fileName}`);
        });
      })
    );

    console.log(
      `Database connected: Host: ${connect.connection.host}, DB Name: ${connect.connection.name}`
    );
  } catch (err) {
    console.log(err);
    process.exit(1);
  }
};

module.exports = { connectDb };
