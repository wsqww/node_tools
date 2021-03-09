// json 转 excel

// const path = require('path');
const fs = require('fs');
const Excel = require("exceljs");

const workbook = new Excel.Workbook();

// 基本的创建信息
// workbook.creator = "Me";
// workbook.lastModifiedBy = "Her";
// workbook.created = new Date();
// workbook.modified = new Date();
// workbook.lastPrinted = new Date();

const worksheet = workbook.addWorksheet("Sheet 1");

// 视图大小， 打开Excel时，整个框的位置，大小
workbook.views = [
  {
    x: 0,
    y: 0,
    width: 1000,
    height: 2000,
    firstSheet: 0,
    activeTab: 1,
    visibility: "visible"
  }
];


const filePath = __dirname + '/before/test.json';
// console.log(filePath); return;
fs.readFile(filePath, 'utf8', function (err, data) {
  if (err) {
    console.log(err);
    return;
  };
  const json = JSON.parse(data); //读取的值
  // console.log(json);
  jsonLoop(json);

  workbook.xlsx.writeFile(__dirname + '/after/test1.xlsx').then(function () {
    console.log('saved');
  });

});


function jsonLoop(json, col = 1, row = 1) {
  let rowLength = 0;
  // console.log('<<<<<<<', rowLength);
  if (Object.keys(json).length > 0) {
    Object.keys(json).forEach((key, index) => {
      // console.log(key, typeof json[key], [col, row + index]);
      const rowStart = row + rowLength;
      const cell = worksheet.getCell(rowStart, col);
      cell.value = key;
      cell.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{argb:'FFFBFF00'},
        // bgColor:{argb:'FFFBFF00'}
      };
      if (typeof json[key] === 'object') {
        const objLength = jsonLoop(json[key], col + 1, rowStart);
        // console.log('length', key, objLength);
        const rowEnd = rowStart + objLength - 1;
        // console.log('$', key, [col, [rowStart, rowEnd]]);
        cell.alignment = { vertical: 'top'};
        worksheet.mergeCells(rowStart, col, rowEnd, col);
        rowLength += objLength;
      } else {
        // console.log('@', key, [col, row + rowLength] );
        worksheet.getCell(rowStart, col+1).value = json[key];
        rowLength++;
      }
    });
  } else {
    rowLength++; // 空 json 需要占一行
  }
  // console.log('>>>>>>>', rowLength);
  return rowLength;
}
