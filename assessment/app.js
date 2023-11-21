const { log } = require("console");
const express = require("express");
const mongoose = require("mongoose");
const fs = require("fs");
const xlsx = require("xlsx");

mongoose.connect("mongodb://127.0.0.1:27017/ADSC");
mongoose.connection.once("open", () => console.log("connected"));

const app = express();
const PORT = 3000;

const Data = mongoose.model("Data", {
  ID: Number,
  " Name": String,
  " Description ": String,
  "Location ": String,
  Price: Number,
  Color: String,
});

const readExcelFile = () => {
  try {
    const workbook = xlsx.readFile("./task.xlsx");

    // Assuming the first sheet is the one you want to read
    const sheetName = workbook.SheetNames[0];

    // Convert the sheet data to JSON
    const excelData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    return excelData;
  } catch (error) {
    console.error("Error reading Excel file:", error);
    return [];
  }
};

app.get("/", async (req, res) => {
  const excelData = readExcelFile();
  try {
    // Insert the data into MongoDB using Mongoose
    const insertedData = await Data.insertMany(excelData);

    res.json({ message: "Data inserted successfully", data: insertedData });
  } catch (error) {
    console.error("Error inserting data into MongoDB:", error);
    res.status(500).json({ error: "Internal Server Error" });
  }
});

app.listen(PORT, () => console.log("server is running"));
