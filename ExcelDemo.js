const ExcelJS = require('exceljs')


async function excelTest(){
 let output = {row:-1, column:-1}   
const workbook = new ExcelJS.Workbook()
await workbook.xlsx.readFile('C:/Users/s_che/ExcelJSUtil/testdownload.xlsx')
const worksheet = workbook.getWorksheet('Sheet1')
worksheet.eachRow((row, rowNumber) => {
                    row.eachCell((cell, colNumber) =>{
                        if (cell.value === 'Banana'){
                          
                           output.row = rowNumber;
                           output.column = colNumber;
                        }
                  })      
                })
   const cell = worksheet.getCell(output.row, output.column)
    cell.value = 'Republic'
    console.log(cell.value)
 workbook.xlsx.writeFile('C:/Users/s_che/ExcelJSUtil/testdownload.xlsx')
            } 

         
excelTest();