const xlsx = require("xlsx");
const zipcodes = require("zipcodes");
const fs = require("fs");
const path = require("path");

const inputFolderPath = "C:\\Users\\itsro\\OneDrive\\Desktop\\delta"; // your folder path
const outputFolderPath = "C:\\Users\\itsro\\OneDrive\\Desktop\\updated delta"; // The path where you want the files to be saved , in this case it will create a folder named updated delta
const referenceZip = "94590"; // to compare with

function extractZipCode(address) {
  if (!address || address.trim() === '') {
    return null; 
  }
    const reversedAddress = address.split('').reverse().join('');
    const zipCodeRegex = /\b\d{5}(?:-\d{4})?\b/;
    const match = reversedAddress.match(zipCodeRegex);
    return match ? match[0].split('').reverse().join('') : null; 
}

function processExcelFile(filePath, outputFilePath, referenceZip) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const jsonData = xlsx.utils.sheet_to_json(sheet);

  const updatedData = jsonData.map((row) => {
    const address = row["address"];
    const zipCode = extractZipCode(address);

    if (zipCode) {
      const distanceInMiles = zipcodes.distance(referenceZip, zipCode);
      const distanceInKm = zipcodes.toKilometers(distanceInMiles);
      row["distance (km)"] = distanceInKm.toFixed(2);
    } else {
      row["distance (km)"] = "N/A";
    }

    return row;
  });

  const updatedSheet = xlsx.utils.json_to_sheet(updatedData);
  const updatedWorkbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(updatedWorkbook, updatedSheet, "Updated Data");

  xlsx.writeFile(updatedWorkbook, outputFilePath);
  console.log(`Updated Excel file saved to ${outputFilePath}`);
}

function processExcelFolder(inputFolderPath, outputFolderPath, referenceZip) {
  fs.readdirSync(inputFolderPath).forEach((file) => {
    const filePath = path.join(inputFolderPath, file);
    if (path.extname(file) === ".xlsx") {
      const outputFilePath = path.join(outputFolderPath, `${file}`);
      processExcelFile(filePath, outputFilePath, referenceZip);
    }
  });
}

if (!fs.existsSync(outputFolderPath)) {
  fs.mkdirSync(outputFolderPath);
}

processExcelFolder(inputFolderPath, outputFolderPath, referenceZip);

const cityName = zipcodes.lookup(referenceZip)
console.warn(`Zip code for city ${cityName.city}`)