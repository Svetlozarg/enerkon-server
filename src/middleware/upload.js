const multer = require("multer");

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, "uploads/");
  },
  filename: function (req, file, cb) {
    if (
      file.originalname === "kcc.xlsx" ||
      file.originalname === "report.docx"
    ) {
      cb(null, file.originalname);
      return;
    }
    const filename = Date.now() + "-" + file.originalname;
    req.customName = filename;
    cb(null, filename);
  },
});

const upload = multer({
  storage: storage,
});

module.exports = upload;
