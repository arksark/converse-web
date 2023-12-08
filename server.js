// server.js

const express = require('express');
const path = require('path');
const fileUpload = require('express-fileupload');
const excelToMongoDb = require('./scripts/excelToMongoDb');
const dateData = require('./scripts/dateData');
const app = express();
app.use(express.static('public'));
const port = 3000;
const PORT = process.env.PORT || 3030;
app.use(fileUpload());

app.post('/upload-excel', (req, res) => {
    if (!req.files || Object.keys(req.files).length === 0) {
        return res.status(400).send('No files were uploaded.');
    }

    const excelFile = req.files.excelFile;
    const filePath = __dirname + '/uploads/' + excelFile.name;

    // Move the file to a location accessible by excelToMongoDb.js
    excelFile.mv(filePath, async (err) => {
        if (err) {
            return res.status(500).send(err);
        }

        // Call the function from excelToMongoDb.js to process the uploaded file
        await excelToMongoDb.processExcelFile(filePath);

        res.json({ success: true });
    });
});

app.post('/run-dateData-script', (req, res) => {
  if (!req.files || Object.keys(req.files).length === 0) {
      return res.status(400).send('No files were uploaded.');
  }

  const excelFile = req.files.excelFile;
  const filePath = __dirname + '/uploads/' + excelFile.name;

  // Move the file to a location accessible by dateData.js
  excelFile.mv(filePath, async (err) => {
      if (err) {
          return res.status(500).send(err);
      }

      // Call the function from dateData.js to process the uploaded file
      await dateData.uploadAndRunDateData();

      res.json({ success: true });
  });
});

app.get('/', (req, res) => {
  // Send the 'index.html' file
  res.sendFile(path.join(__dirname, 'public' , 'index.html'));
});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
