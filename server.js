import express from 'express';
import cors from 'cors';
import bodyParser from 'body-parser';
import fs from 'fs';
import XLSX from 'xlsx';

const app = express();
const PORT = process.env.PORT || 3000;
const EXCEL_FILE = 'cghs_correction.xlsx';

app.use(cors());
app.use(bodyParser.json());

app.post('/submit', (req, res) => {
  const formData = req.body;
  let workbook, worksheet;

  // If file exists, read it; else, create new workbook
  if (fs.existsSync(EXCEL_FILE)) {
    workbook = XLSX.readFile(EXCEL_FILE);
    worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(worksheet);
    data.push(formData);
    const newWs = XLSX.utils.json_to_sheet(data);
    workbook.Sheets[workbook.SheetNames[0]] = newWs;
  } else {
    const ws = XLSX.utils.json_to_sheet([formData]);
    workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, ws, 'Sheet1');
  }

  XLSX.writeFile(workbook, EXCEL_FILE);
  res.json({ success: true });
});

app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
}); 