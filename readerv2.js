const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const inputFolder = './input';
const outputFolder = './output';
const outputFileName = 'output.xlsx';

const workbook = new ExcelJS.Workbook();
const outputWorkbook = new ExcelJS.Workbook();

fs.readdir(inputFolder, (err, files) => {
  if (err) {
    console.error(err);
    return;
  }

  files.forEach((file) => {
    if (path.extname(file) === '.xlsx') {
      workbook.xlsx.readFile(path.join(inputFolder, file), (err) => {
        if (err) {
          console.error(err);
          return;
        }

        const worksheet = workbook.getWorksheet(1);
        const questions = [];

        worksheet.eachRow((row) => {
          questions.push(row.getCell('A').value);
        });

        const outputWorksheet = outputWorkbook.addWorksheet(file);
        let rowNumber = 1;

        questions.forEach((question) => {
          outputWorksheet.getCell(`A${rowNumber}`).value = question;
          rowNumber++;
        });

        outputWorkbook.xlsx.writeFile(path.join(outputFolder, outputFileName), (err) => {
          if (err) {
            console.error(err);
            return;
          }
        });
      });
    }
  });
});
