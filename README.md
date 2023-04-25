# Excel Attachment Finder
## Description
This is a Node.js project that reads an Excel file from a specified folder, finds the attachment files corresponding to each row of the Excel file, and adds the file paths of the attachments to a new column in the Excel file. Once the process is complete, the Excel file is moved to a different folder.

## Prerequisites
To run this project, you will need to have Node.js installed on your computer. You can download it from the [official website](https://nodejs.org/en/download/)

## Installation
1. Clone this repository to your local machine.
2. Open a terminal or command prompt and navigate to the project directory.
3. Run the command **`npm install`** to install the project dependencies.

## Usage
1. Place the Excel file to be processed in the **`raw_files`** folder.
2. Place the attachment files in the **`attachments`** folder with filenames matching the unique identifiers in the Excel file.
3. Open a terminal or command prompt and navigate to the project directory.
4. Run the command **`node index.js`** to start the application.
5. The application will read the Excel file and find the attachment files. The file paths of the attachments will be added to a new column in the Excel file.
6. Once the process is complete, the Excel file will be moved to the **`completed_files`** folder.

## Dependencies
This project requires the following dependencies:

+ **`xlsx`**: A library for reading and writing Excel files.
+ **`path`**: A built-in Node.js module for working with file and directory paths.
+ **`moment`**: A library for working with dates and times.
+ **`fs`**: A built-in Node.js module for working with the file system.

## Support
If you have any questions or issues with this project, please contact **`Ankur Prajapati`** at **`ankur.b@acldigital.com`**.

Feel free to modify this README file to suit your specific project needs.
