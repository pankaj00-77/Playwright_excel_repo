const Exceljs = require('exceljs');

async function excelTest() {

    let output = {row:-1,col:-1};// with the help of this we can access them from any block by make global variable

    const workbook = new Exceljs.Workbook();//This line is used to create a new Excel workbook (just like opening a new Excel file in Microsoft Excel).
    await workbook.xlsx.readFile("C:/Users/Pankaj/Downloads/ExceldownloadTest.xlsx");//"Go to this file on my computer and open the Excel file so I can read data from it."
    const worksheet = workbook.getWorksheet('Sheet1');//"Get the worksheet (tab) named 'Sheet1' from the Excel file (workbook)."
    //"Go through every row in the sheet, and then go through every cell in each row, and print the value of each cell."
    worksheet.eachRow((row, rowNumber) =>//Goes through each row in the Excel sheet.  For each row, this function runs. rowNumber gives row index (like 1, 2, 3...)
    {
        row.eachCell((cell, cellNumber) =>//Now, go through each cell in the current row.  For each cell in the row, this function runs.
        {
            if (cell.value === "Banana") {// it gives the index of the apple in the excel table 
                console.log(output.row = rowNumber);//console.log(cell.rowNumber); both are same 
                console.log(output.col = cellNumber);//console.log(cell.columnNumber); both are work same
            }
            // console.log(cell.value); // Print the value inside the current cell (like "Rajpal", "test@gmail.com", etc.)
        }
        );

    });
    const cell  = worksheet.getCell(output.row,output.col);// locate the cell
    cell.value = "Wine";//change the value of that cell 
    await workbook.xlsx.writeFile("C:/Users/Pankaj/Downloads/ExceldownloadTest.xlsx");// but by executing this line we will save the changes the excel file by using the writeFile function


}
excelTest();// calling that function 
// to run the excel file or code write----------->node <file name.js> like----> node excelDemo.js and it prints the value 


