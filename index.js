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
      
      let i = 1,j = 1,k = 1 ,l = 1;
      let billInitial = '',expInitial = '',depInitial = '',jeInitial = '';
      for (const row of jsonData) {
        var billFolderName = [];
        
        if (row["Transaction Type"] == "Bill") {
          const billAmount = row["Credit"].toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
          billFolderName = ['attachments/Bill/Bill ' + row["Name"] + ' \u20B9' + billAmount + '.00','attachments/Bill/Bill #' + row["No."] + ' ' + row["Name"] + ' \u20B9' + billAmount + '.00','attachments/Bill/Bill ' + row["No."] + ' ' + row["Name"] + ' ' + billAmount + '.00'];
          if(billInitial !== "Bill_"+moment(row["Date"],'DD/MM/YYYY').format('MMM_YY')){
            i = 1;
          }
          row["__EMPTY"] = "Bill_"+moment(row["Date"],'DD/MM/YYYY').format('MMM_YY')+"_"+i;
          i++;
          billInitial = "Bill_"+moment(row["Date"],'DD/MM/YYYY').format('MMM_YY');
        }
        else if(row["Transaction Type"] == "Deposit"){
          const depAmount = row["Debit"].toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
          billFolderName = ['attachments/Deposit/Deposit \u20B9' + depAmount + '.00','attachments/Deposit/Deposit ' + row["Name"] + ' \u20B9' + depAmount + '.00'];
          if(depInitial !== "Dep_"+moment(row["Date"],'DD/MM/YYYY').format('MMM_YY')){
            j = 1;
          }
          row["__EMPTY"] = "Dep_"+moment(row["Date"],'DD/MM/YYYY').format('MMM_YY')+"_"+j;
          j++;
          depInitial = "Dep_"+moment(row["Date"],'DD/MM/YYYY').format('MMM_YY');
        }
        else if(row["Transaction Type"] == "Expense"){
          const expAmount = row["Credit"].toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
          billFolderName = ['attachments/Expense/Expense \u20B9' + expAmount + '.00','attachments/Expense/Expense \u20B9' + expAmount + '.00-1','attachments/Expense/Expense '+ row["Name"] +' \u20B9' + expAmount + '.00','attachments/Expense/Expense '+ row["Name"] +' \u20B9' + expAmount + '.00-1'];
          if(expInitial !== "Exp_"+moment(row["Date"],'DD/MM/YYYY').format('MMM_YY')){
            k = 1;
          }
          row["__EMPTY"] = "Exp_"+moment(row["Date"],'DD/MM/YYYY').format('MMM_YY')+"_"+k;
          k++;
          expInitial = "Exp_"+moment(row["Date"],'DD/MM/YYYY').format('MMM_YY');
        }
        else if(row["Transaction Type"] == "Journal Entry"){
          if(jeInitial !== "JE_"+moment(row["Date"],'DD/MM/YYYY').format('MMM_YY')){
            l = 1;
          }
          row["__EMPTY"] = "JE_"+moment(row["Date"],'DD/MM/YYYY').format('MMM_YY')+"_"+l;
          l++;
          jeInitial = "JE_"+moment(row["Date"],'DD/MM/YYYY').format('MMM_YY');
        }
        
        const folderName = await findExistingFolder(billFolderName);
        if (folderName) {
          let foundFilePath;
          let destinationPath;
          if(row["Transaction Type"] == "Expense"){
            const fileWordCombo = "_"+row["Credit"]+"_"+moment(row["Date"],'DD/MM/YYYY').format('DDMM');
            foundFilePath = await findFilePathByWordCombination(folderName,fileWordCombo);
          }
          else{
            const attachFile = await fs.promises.readdir(folderName);
            foundFilePath = attachFile.length > 0 ? path.join(folderName, attachFile[0]) : "NA";
          }

          if(foundFilePath !== "NA"){
            const fileExt = path.extname(foundFilePath);
            const destFolder= "processed_files/"+moment(row["Date"],'DD/MM/YYYY').format('MMM_YY');
            
            // Get the directory of the destination file
            // const sourcePath = path.join(__dirname,foundFilePath);
            destinationPath = path.join(destFolder,row["__EMPTY"]+fileExt);
            const destFolderName = path.join(__dirname,destFolder);

            // Check if the directory exists, and create it if it doesn't
            if (!fs.existsSync(destFolder)) {
              fs.mkdirSync(destFolder, { recursive: true });
            }

            // Copy the file to the destination
            fs.copyFileSync(foundFilePath, destinationPath);
            
            row["URL"] = foundFilePath;
            row["New URL"] = destinationPath;
          }
          else{
            row["URL"] = "NA";
            row["New URL"] = "NA";
          }
        } else {
          row["URL"] = "NA";          
          row["New URL"] = "NA";          
        }
      }

      const oldKey = "__EMPTY";
      const newKey = "External Id";
      const newArr = jsonData.map(obj => {
        const newObj = {};
        Object.keys(obj).forEach(key => {
          if (key === oldKey) {
            newObj[newKey] = obj[key];
          } else {
            newObj[key] = obj[key];
          }
        });
        return newObj;
      });
      // console.log(jsonData);

      const headers = Object.keys(newArr[0]);

      const data = newArr.map(row => {
        return headers.map(header => row[header]);
      });

      // const newSheet = XLSX.utils.json_to_sheet(newArr);
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

