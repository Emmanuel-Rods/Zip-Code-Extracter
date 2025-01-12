const xlsx = require("xlsx");
const zipcodes = require("zipcodes");
const fs = require("fs");
const path = require("path");

const inputFolderPath =
  "C:\\Users\\itsro\\OneDrive\\Desktop\\Unity Township (Pennsylvania)"; // your folder path
const outputFolderPath = "C:\\Users\\itsro\\OneDrive\\Desktop\\updated unity;"; // The path where you want the files to be saved , in this case it will create a folder named updated delta
const referenceZip = "15601"; // to compare with
const maxDistanceInKm = 30; // Maximum distance in kilometers to retain rows

function extractZipCode(address) {
  if (!address || address.trim() === "") {
    return null;
  }
  const reversedAddress = address.split("").reverse().join("");
  const zipCodeRegex = /\b\d{5}(?:-\d{4})?\b/;
  const match = reversedAddress.match(zipCodeRegex);
  return match ? match[0].split("").reverse().join("") : null;
}

function processExcelFile(filePath, outputFilePath, referenceZip) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const jsonData = xlsx.utils.sheet_to_json(sheet);

  const filteredData = jsonData.filter((row) => {
    const address = row["address"];
    if (!address) {
      row["distance (km)"] = "N/A"; // Mark as N/A
      return true; // Keep rows with no address
    }

    const zipCode = extractZipCode(address);
    if (!zipCode) {
      row["distance (km)"] = "N/A"; // Mark as N/A
      return true; // Keep rows with no valid zip code
    }
   
    if (zipCode === referenceZip) {
      row["distance (km)"] = "0.00"; // Distance is zero
      return true; // Keep the row
    }

    const distanceInMiles = zipcodes.distance(referenceZip, zipCode);
    if (distanceInMiles === undefined || distanceInMiles === null) {
      row["distance (km)"] = "N/A"; // Mark as N/A
      return true; // Keep rows with invalid distances
    }

    const distanceInKm = zipcodes.toKilometers(distanceInMiles);
    if (distanceInKm <= maxDistanceInKm) {
      row["distance (km)"] = distanceInKm.toFixed(2); // Add the distance for valid rows
      return true; // Keep rows within the max distance
    }

    // Exclude rows that exceed max distance
    return false;
  });

  const updatedSheet = xlsx.utils.json_to_sheet(filteredData);
  const updatedWorkbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(updatedWorkbook, updatedSheet, "Filtered Data");

  xlsx.writeFile(updatedWorkbook, outputFilePath);
  console.log(`Filtered Excel file saved to ${outputFilePath}`);
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