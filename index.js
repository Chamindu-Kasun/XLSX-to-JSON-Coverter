const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const path = require("path");
const dotenv = require("dotenv");
dotenv.config();
const cors = require("cors");

const app = express();
app.use(cors({ origin: "*" }));

app.use(express.static(`${__dirname}/public`));

app.get("/", (req, res) => {
  res.sendFile(path.join(`${__dirname}/public`, "index.html"));
});

const storage = multer.memoryStorage();
const upload = multer({
  storage,
  limits: {
    fileSize: 5 * 1024 * 1024,
  },
  fileFilter: (req, file, cb) => {
    if (
      file.mimetype !==
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ) {
      return cb(new Error("Only XLSX files are allowed"));
    }
    cb(null, true);
  },
});

app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send("No file uploaded");
    }
    const buffer = req.file.buffer;
    const workbook = xlsx.read(buffer, { type: "buffer" });

    const sheetNames = workbook.SheetNames;
    const firstSheetName = sheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = xlsx.utils.sheet_to_json(worksheet);
    const jsonString = JSON.stringify(jsonData, null, 2);
    res.setHeader("Content-Disposition", "attachment; filename=output.json");
    res.setHeader("Content-Type", "application/json");
    res.send(jsonString);
  } catch (error) {
    console.error("Error processing file:", error.message);
    res.status(500).send("Error processing file");
  } finally {
    if (req.file) {
      req.file.buffer = null;
    }
  }
});

app.use((req, res, next) => {
  if (req.url !== "/upload") {
    return res.sendFile(path.join(`${__dirname}/public`, "index.html"));
  }
  if (req.method === "GET" && req.url === "/upload") {
    return res.sendFile(path.join(`${__dirname}/public`, "index.html"));
  }
  next();
});

app.listen(process.env.PORT, () => {
  console.log(`Server started on port ${process.env.PORT}`);
});
