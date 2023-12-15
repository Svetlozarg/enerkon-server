const mongoose = require("mongoose");
const fs = require("fs");
const path = require("path");
let Buffer = require("buffer").Buffer;
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

    const uploadsFolderPath = path.join("./uploads");

    if (!fs.existsSync(uploadsFolderPath)) {
      fs.mkdirSync(uploadsFolderPath);
      console.log("Uploads folder created.");
    } else {
      console.log("Uploads folder already exists.");
    }

    for (const fileUrl of FILE_URLS) {
      const response = await fetch(fileUrl);
      const arrayBuffer = await response.arrayBuffer();

      const buffer = Buffer.from
        ? Buffer.from(arrayBuffer)
        : new Buffer(arrayBuffer);
      // a

      const fileName = path.basename(fileUrl);

      const filePath = path.join(uploadsFolderPath, fileName);
      fs.writeFileSync(filePath, buffer, (err) => {
        if (err) throw err;
        console.log(`File downloaded and saved: ${fileName}`);
      });
    }

    console.log(
      "Database connected: ",
      "Host:" + connect.connection.host,
      "DB Name:" + connect.connection.name
    );
  } catch (err) {
    console.log(err);
    process.exit(1);
  }
};

module.exports = { connectDb };
