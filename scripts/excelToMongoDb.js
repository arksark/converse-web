const { MongoClient } = require('mongodb');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

const dbName = 'excelDb';
const collectionName = 'users';

// Function to generate a random password
const generateRandomPassword = () => {
  const length = 8;
  const charset = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  let password = "";
  for (let i = 0; i < length; i++) {
    const randomIndex = Math.floor(Math.random() * charset.length);
    password += charset.charAt(randomIndex);
  }
  return password;
};

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

// Function to fetch the password from the database based on ID
async function getPasswordFromDatabase(client, id) {
  try {
    const db = client.db(dbName);
    const collection = db.collection(collectionName);
    const result = await collection.findOne({ id });
    return result ? result.password : undefined;
  } catch (error) {
    console.error('Error fetching password from the database:', error);
    throw error;
  }
}

// Function to insert or update user data in the database
async function insertOrUpdateUserData(client, id, name, department) {
  try {
    const db = client.db(dbName);
    const collection = db.collection(collectionName);

    const existingPassword = await getPasswordFromDatabase(client, id);
    const password = existingPassword !== undefined ? existingPassword : generateRandomPassword();

    const existingDocument = await collection.findOne({ id });

    if (!existingDocument) {
      // If the document with the ID doesn't exist, insert a new document
      const document = { id, name, department, password };
      await collection.insertOne(document);
      console.log('User Document inserted:', document);
    } else {
      // If the document with the ID exists, update it with new fields if provided
      const updateFields = {};
      
      if (name) updateFields.name = name;
      if (department) updateFields.department = department;

      // Using updateOne with $set to update only the specified fields
      await collection.updateOne({ id }, { $set: updateFields });
      console.log('User Document updated with new fields:', updateFields);
    }
  } catch (error) {
    console.error('Error inserting or updating user data:', error);
    throw error;
  }
}

// Function to process employee data from the Excel sheet
async function processEmployeeData(client, data) {
  try {
    // Cleaning data by removing undefined and null cells
    const cleanedData = data.map(row => row.filter(cell => cell !== undefined && cell !== null));

    // Filtering rows with seven-digit IDs
    const rowsWithSevenDigitIDs = cleanedData.filter(row => row[0] !== null && /^\d{7}$/.test(String(row[0])));

    await Promise.all(rowsWithSevenDigitIDs.map(async row => {
      const id = String(row[0]);
      const name = row[1] !== null ? String(row[1]) : '';
      const department = row[2] !== null ? String(row[2]) : '';

      // Call the function to insert or update user data in the database
      await insertOrUpdateUserData(client, id, name, department);
    }));

    // Replace the line below with the call to fetch and insert date data
    // await fetchAndInsertDateData(client, documentsToInsert);
  } catch (error) {
    console.error('Error processing employee data:', error);
    throw error;
  }
}

// Function to process all Excel files in a folder
async function processAllExcelFiles(client, excelFilePath) {
  try {
      const workbook = XLSX.readFile(excelFilePath);
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      const data = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

      // Cleaning and processing data...
      await processEmployeeData(client, data);
  } catch (error) {
      console.error('Error processing all Excel files:', error);
      throw error;
  }
}

// Function to process the Excel file
async function processExcelFile(filePath) {
  let client;

  try {
      // Establish MongoDB connection
      client = await connectToMongoDB();

      // Read the Excel file
      const workbook = XLSX.readFile(filePath);
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      const data = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

      // Process employee data
      await processEmployeeData(client, data);
  } catch (error) {
      console.error('Error processing Excel file:', error);
  } finally {
      // Close MongoDB connection
      if (client) {
          await closeMongoDBConnection(client);
      }
  }
}

module.exports = {
  processExcelFile,
};