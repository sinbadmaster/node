/*
 * @Author: sinbadMaster
 * @Description: 将json转为excel
 * @Date: 2020-06-09 08:59:11
 * @LastEditTime: 2020-06-17 10:02:31
 * @FilePath: \json2Excel\index.js
 */ 

const Excel = require('exceljs');
const path = require('path');

/**
 * @description: 将对象或数组对象转为excel, 暂时只支持了数组嵌套对象，不支持多重嵌套
 * @param {string} savePath 保存路径
 * @param {string} fileName 文件名 
 * @param {Object} jsonData json格式的数据
 * @return: 
 * @Date: 2020-06-09 15:13:30
 */
exports.default = function json2Excel(savePath, fileName, jsonData) {
  savePath = savePath.endsWith(path.sep) ? savePath : savePath + path.sep;
  fileName = (fileName.endsWith('.xls') || fileName.endsWith('.xlsx'))
    ? fileName
    : fileName + '.xlsx'

  const workbook = new Excel.stream.xlsx.WorkbookWriter({
    filename: `${savePath}${fileName}`
  });

  const worksheet = workbook.addWorksheet('sheet');

  const isArray = Array.isArray(jsonData);
  const header = isArray
    ? jsonData.length
        ? jsonData[0]
        : {}
    : jsonData;
  
  worksheet.columns = Object.keys(header).map(prop => ({ header: prop, key: prop }));

  if (isArray) {
    for (const row of jsonData) {
      worksheet.addRow(row).commit();
    }
  } else {
    worksheet.addRow(jsonData).commit();
  }
  
  workbook.commit();
}
