/*
 * @Author: sinbadMaster
 * @Description: 将json转为excel
 * @Date: 2020-06-09 08:59:11
 * @LastEditTime: 2020-06-17 14:29:41
 * @FilePath: \json2Excel\index.js
 */ 

const Excel = require('exceljs');
const path = require('path');
const fs = require('fs');

/**
 * @description: 将对象或数组对象转为excel, 暂时只支持了数组嵌套对象，不支持多重嵌套
 * @param {string} savePath 保存路径
 * @param {string} fileName 文件名 
 * @param {Object} jsonData json格式的数据
 * @param {{headers: {[string]: string}}} options 选项参数可选，如显示的传入headers则以此为表头，
 * 如未传入则以数组第一项对象的key值作为表头
 * @return: 
 * @Date: 2020-06-09 15:13:30
 */
module.exports = function json2Excel(savePath, fileName, jsonData, options) {
  savePath = savePath.endsWith(path.sep) ? savePath : savePath + path.sep;
  fileName = (fileName.endsWith('.xls') || fileName.endsWith('.xlsx'))
    ? fileName
    : fileName + '.xlsx'

  if (!fs.existsSync(savePath)) {
    fs.mkdirSync(savePath, { recursive: true })
  }
  
  const workbook = new Excel.stream.xlsx.WorkbookWriter({
    filename: `${savePath}${fileName}`
  });

  const worksheet = workbook.addWorksheet('sheet');

  const isArray = Array.isArray(jsonData);
  
  if (typeof options === 'object' && typeof options.headers === 'object') {
    worksheet.columns = Object.entries(options.headers).map(item => ({header: item[1], key: item[0]}));
  } else {
    const header = isArray
      ? jsonData.length
          ? jsonData[0]
          : {}
      : jsonData;
    
    worksheet.columns = Object.keys(header).map(prop => ({ header: prop, key: prop }));
  }  

  if (isArray) {
    for (const row of jsonData) {
      worksheet.addRow(row).commit();
    }
  } else {
    worksheet.addRow(jsonData).commit();
  }
  
  workbook.commit();
}
