const express = require("express");
const { resolve } = require("path");
const Excel = require("exceljs");
const path = require("path");

const app = express();
const port = 3010;

app.use(express.static("static"));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

async function excelBuilder(columns, data) {
  const newWorkbook = new Excel.Workbook();
  const worksheet = newWorkbook.addWorksheet("My Sheet");

  const updatedColumns = columns.map((col) => {
    col.header = col.header.map((h) => h.text);
    col.width = 10;
    return col;
  });

  worksheet.columns = updatedColumns;

  const flatArray = data.map((d) => {
    return Object.keys(d)
      .filter((key) => !isNaN(key))
      .map((key) => d[key]);
  });
  worksheet.addRows(flatArray);

  const timestamp = new Date().toISOString().replace(/[-:T.]/g, "");
  const fileName = `export_${timestamp}.xlsx`;
  const filePath = path.resolve(__dirname, "exports", fileName);

  await newWorkbook.xlsx.writeFile(filePath);

  return filePath;
}

app.get("/", async (req, res) => {
  const { columns, data } = req.body;

  if (req.body) {
    res.status(500).send("Must to pass body request");
  }

  await res.download(await excelBuilder(columns, data), (err) => {
    if (err) {
      res.status(500).send("Error sending file");
    }
  });
});

app.listen(port, () => {
  console.log(`Example app listening at http://localhost:${port}`);
});
