const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const moment = require('moment');

const folderPath = 'raw_files'; // Specify the folder path
const completedPath = 'completed_files'; // Specify the completed folder path

async function processFiles() {
  const files = await fs.promises.readdir(folderPath);
  if (files.length) {
    for (const file of files) {
      const filePath = path.join(folderPath, file);
      const workbook = XLSX.readFile(filePath);
      const filename = path.basename(filePath);
      const datetime = new Date().toISOString().replace(/[-T:\.Z]/g, '');
      const uniqueFilename = `${datetime}-${filename}`;
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const options = {range:4,defval:''};
      const jsonData = XLSX.utils.sheet_to_json(worksheet,options);

      for (const row of jsonData) {
        delete row["__EMPTY"];
        var billFolderName = [];
        if (row["Transaction Type"] == "Bill") {
          const billAmount = row["Credit"].toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
          billFolderName = ['attachments/Bill/Bill ' + row["Name"] + ' \u20B9' + billAmount + '.00','attachments/Bill/Bill #' + row["No."] + ' ' + row["Name"] + ' \u20B9' + billAmount + '.00','attachments/Bill/Bill ' + row["No."] + ' ' + row["Name"] + ' ' + billAmount + '.00'];
          // const attachmentPath = 'attachments/Bill/' + billFolderName ;

          // console.log(billFolderName);
        }
        else if(row["Transaction Type"] == "Deposit"){
          // console.log(row);
          const depAmount = row["Debit"].toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
          billFolderName = ['attachments/Deposit/Deposit \u20B9' + depAmount + '.00','attachments/Deposit/Deposit ' + row["Name"] + ' \u20B9' + depAmount + '.00'];
        }
        else if(row["Transaction Type"] == "Expense"){
          // console.log(row);
          const expAmount = row["Credit"].toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
          billFolderName = ['attachments/Expense/Expense \u20B9' + expAmount + '.00','attachments/Expense/Expense \u20B9' + expAmount + '.00-1'];
        }
        
        const folderName = await findExistingFolder(billFolderName);
        // console.log(folderName);
        if (folderName) {
          let foundFilePath;
          if(row["Transaction Type"] == "Expense"){
            const fileWordCombo = "_"+row["Credit"]+"_"+moment(row["Date"],'DD/MM/YYYY').format('DDMM');
            // console.log(row["Date"]+" converted date is "+moment(row["Date"],'DD/MM/YYYY').format('DDMM'));
            // console.log(fileWordCombo);
            foundFilePath = await findFilePathByWordCombination(folderName,fileWordCombo);
            // console.log(foundFilePath);
          }
          else{
            const attachFile = await fs.promises.readdir(folderName);
            foundFilePath = attachFile.length > 0 ? path.join(folderName, attachFile[0]) : "NA";
            // row["URL"] = foundFilePath;
            // console.log(foundFilePath);
          }
          row["URL"] = foundFilePath;
        } else {
          row["URL"] = "NA";
        }
        // writing condition for transaction type deposit
      }

      // console.log(jsonData);
      const headers = Object.keys(jsonData[0]);

      const data = jsonData.map(row => {
        return headers.map(header => row[header]);
      });

      // const newSheet = XLSX.utils.json_to_sheet(jsonData);
      const newSheet = XLSX.utils.aoa_to_sheet([headers, ...data]);

      const newWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWorkbook, newSheet, sheetName);
      // await fs.promises.mkdir(completedPath, { recursive: true });
      await XLSX.writeFile(newWorkbook, path.join(completedPath, uniqueFilename));
    }
  } else {
    console.log(files.length);
  }
}

// async function isFolderEmptyOrNotExists(folderPath) {
//   try {
//     const files = await fs.promises.readdir(folderPath);
//     return files.length === 0;
//   } catch (err) {
//     return true;
//   }
// }

async function findExistingFolder(folderNames) {
  let folderExists = '';
  for (const folderName of folderNames) {
    const folderAttachmentPath = path.join(__dirname, folderName);
    try {
      const files = await fs.promises.readdir(folderAttachmentPath);
      if (files.length > 0) {
        folderExists = folderName;
      }
      else{
        console.log("No file")
        // folderExists = "NA";
      }
    } catch (err) {
      console.log("No folder");
      // folderExists = "NA";
    }
  }
  return folderExists;
}

async function findFilePathByWordCombination(folderName, wordCombination) {
  let fileExists;
  try {
    const attachFile = await fs.promises.readdir(folderName);
    const foundFileName = attachFile.find(fileName => fileName.includes(wordCombination));
    if (foundFileName) {
      fileExists = path.join(folderName, foundFileName);
    } else {
      fileExists = "NA";
    }
  } catch (err) {
    // Handle error if needed
    fileExists = "NA";
  }
  return fileExists;
}


processFiles().then(() => {
  console.log("All files processed successfully");
}).catch((err) => {
  console.error(err);
});

