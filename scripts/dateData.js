// dateData.js

const { MongoClient } = require('mongodb');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const folderPath = 'F:\\Arief-IT Intern\\Dokumen\\'; // Replace with the actual path to your folder
const dbName = 'excelDb'; // Replace 'excelDb' with your actual database name
const collectionName = 'users'; // Replace 'users' with your actual collection name

// Function to establish MongoDB connection
async function connectToMongoDB() {
  try {
    const client = new MongoClient('mongodb://localhost:27017');
    await client.connect();
    console.log('Connected to MongoDB');
    return client;
  } catch (error) {
    console.error('Error connecting to MongoDB:', error);
    throw error;
  }
}

// Function to close MongoDB connection
async function closeMongoDBConnection(client) {
  try {
    await client.close();
    console.log('MongoDB connection closed');
  } catch (error) {
    console.error('Error closing MongoDB connection:', error);
  }
}

// Function to get the newest Excel file in the folder
function getNewestExcelFile(folderPath) {
  const files = fs.readdirSync(folderPath);
  const excelFiles = files.filter(file => file.endsWith('.xls') || file.endsWith('.xlsx'));

  if (excelFiles.length === 0) {
    console.error('No Excel files found in the folder.');
    return null;
  }

  // Get the file with the latest modification time
  const newestFile = excelFiles.reduce((prevFile, currentFile) => {
    const prevFilePath = path.join(folderPath, prevFile);
    const currentFilePath = path.join(folderPath, currentFile);

    const prevFileStats = fs.statSync(prevFilePath);
    const currentFileStats = fs.statSync(currentFilePath);

    return prevFileStats.mtimeMs > currentFileStats.mtimeMs ? prevFile : currentFile;
  });

  return newestFile;
}

// Function to insert or update data in the database
async function insertOrUpdateData(client, id, newDates, newTimeIn, newTimeOut) {
  try {
    const db = client.db(dbName);
    const collection = db.collection(collectionName);

    // Find the document with the given ID
    const existingDoc = await collection.findOne({ id });

    if (existingDoc) {
      // If the document with the ID exists, clean the existing date entries
      await collection.updateOne(
        { id },
        {
          $set: {
            dates: [],
            timeIn: [],
            timeOut: [],
          },
        }
      );

      // Add new date entries
      await collection.updateOne(
        { id },
        {
          $push: {
            dates: { $each: newDates },
            timeIn: { $each: newTimeIn },
            timeOut: { $each: newTimeOut },
          },
        }
      );

      console.log(`Data cleaned and updated in the database for ID: ${id}`);
    } else {
      // If the document with the ID doesn't exist, insert a new document
      await collection.insertOne({
        id,
        dates: newDates,
        timeIn: newTimeIn,
        timeOut: newTimeOut,
      });
      console.log(`Data inserted into the database for ID: ${id}`);
    }
  } catch (error) {
    console.error('Error inserting/updating data into the database:', error);
  }
}

// Function to process rows from the Excel sheet and insert/update data in the database
async function processRowsAndInsertData(sheet, client, filePath) {
  try {
    console.log(`Processing file: ${filePath}`);
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

    // Fetch all existing IDs and date/time data in the database
    const db = client.db(dbName);
    const collection = db.collection(collectionName);
    const existingData = await collection.find({}, { projection: { id: 1, dates: 1, timeIn: 1, timeOut: 1 } }).toArray();

    // Process each row in the Excel file
    for (let i = 0; i < rows.length; i++) {
      const id = String(rows[i][0]);

      if (/^\d{7}$/.test(id)) {
        // If the ID is a seven-digit number
        const dates = [];
        const timeIn = [];
        const timeOut = [];

        // Fetch dates starting from 2 rows below until an undefined or empty cell is encountered
        let rowIndex = i + 2;
        while (rows[rowIndex] && rows[rowIndex][0] !== undefined && rows[rowIndex][0] !== '') {
          const date = String(rows[rowIndex][0]);
          dates.push(date);

          const checkIn = rows[rowIndex][2] !== undefined ? String(rows[rowIndex][2]) : '0';
          const checkOut = rows[rowIndex][3] !== undefined ? String(rows[rowIndex][3]) : '0';

          timeIn.push(checkIn);
          timeOut.push(checkOut);

          rowIndex++;
        }

        // Call the function to insert/update data in the database
        await insertOrUpdateData(client, id, dates, timeIn, timeOut);

        // Remove the ID from the existingData array (IDs with data in the Excel file)
        const index = existingData.findIndex((data) => data.id === id);
        if (index !== -1) {
          existingData.splice(index, 1);
        }
      }
    }

    // Clean up date/time data for IDs not found in the Excel file
    await Promise.all(
      existingData.map(async ({ id, dates, timeIn, timeOut }) => {
        // Clean up date/time data for the ID
        await collection.updateOne(
          { id },
          {
            $set: {
              dates: [],
              timeIn: [],
              timeOut: [],
            },
          }
        );
        console.log(`Data cleaned for ID not found in the Excel file: ${id}`);
      })
    );
  } catch (error) {
    console.error('Error processing rows and inserting/updating data:', error);
    throw error;
  }
}

// Export a function that can be called when the button is pressed
async function uploadAndRunDateData(filePath) {
  let client;

  try {
    // Establish MongoDB connection
    client = await connectToMongoDB();

    // Get the newest Excel file in the folder
    const newestFile = getNewestExcelFile(folderPath);

    if (newestFile) {
      const filePath = path.join(folderPath, newestFile);
      console.log(`Processing file: ${filePath}`);

      // Read the Excel file and process the data
      const workbook = XLSX.readFile(filePath);
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      // Call the function to process rows and insert/update data in the database
      await processRowsAndInsertData(sheet, client);
    }
  } catch (error) {
    console.error('Error:', error);
  } finally {
    // Close MongoDB connection
    if (client) {
      await closeMongoDBConnection(client);
    }
  }
}

// Export other functions as needed
module.exports = {
  uploadAndRunDateData,
};
