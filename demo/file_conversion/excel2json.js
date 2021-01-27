const fs = require('fs');
const Excel = require("exceljs");

let jsonData = {};

// return;

const workbook = new Excel.Workbook();
workbook.xlsx.readFile(__dirname + '/before/en.xlsx').then(() => {
  const worksheet = workbook.getWorksheet(1);
  worksheet.eachRow((row, rowNum) => {
    const values = row.values;
    const keys = values.filter((key, index) => index > 0);
    jsonDeep(jsonData, keys);
  });

  // console.log(jsonData);
  // return;
  const jsonStr = JSON.stringify(jsonData, null, '\t');
  // console.log('>>>>', jsonStr);
  // return;
  fs.writeFile(__dirname + '/after/en1.json', jsonStr, function (err) {
    if (err) {
      console.log(err);
      return;
    };
    console.log('saved');
  })
});


function jsonDeep(obj, keys) {
  if (keys.length > 2) {
    // console.log(keys);
    if (!obj[keys[0]]) {
      obj[keys[0]] = {};
    }
    const newKeys = keys.filter((key, index) => index > 0);
    jsonDeep(obj[keys[0]], newKeys);
  }
  if (keys.length === 2) {
    if (!obj[keys[0]]) {
      obj[keys[0]] = keys[1];
    } else {
      obj[keys[0]][keys[1]] = {};
    }
  }
}
