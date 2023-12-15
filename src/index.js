const express = require("express");
const cors = require("cors");
const errorHandler = require("./middleware/errorHandler");
const { connectDb } = require("./config/dbConnection");
require("dotenv").config();

connectDb();
const app = express();

const port = process.env.PORT || 5000;

app.use(express.json());
app.use(cors());
app.use("/api", require("./routes/authRoutes"));
app.use("/api/users", require("./routes/userRoutes"));
app.use("/api/project", require("./routes/projectRoutes"));
app.use("/api/document", require("./routes/documentRoutes"));
app.use("/uploads", express.static("uploads"));
app.use(errorHandler);
app.use((req, res) => {
  res.status(404).json({ error: "Route not found" });
});

app.listen(port, () => {
  console.log(
    `Server successfully started and running on http://localhost:${port}`
  );
});
