const Exceljs = require('exceljs');

async function writeExcelTest(searchText, replaceText, pathFile) { //here we passes some parameters so we don't need to make change everywhere everytime we can call them while calling the function

    const workbook = new Exceljs.Workbook(); //This line is used to create a new Excel workbook (just like opening a new Excel file in Microsoft Excel).
    await workbook.xlsx.readFile(pathFile); //"Go to this file on my computer and open the Excel file so I can read data from it."
    const worksheet = workbook.getWorksheet('Sheet1'); //"Get the worksheet (tab) named 'Sheet1' from the Excel file (workbook)."
    
    //"Go through every row in the sheet, and then go through every cell in each row, and print the value of each cell."
    const output = await readExcel(worksheet, searchText); //call function with worksheet and searchText parameter
    
    if (output.row !== -1 && output.col !== -1) { //  added this condition to avoid invalid row/column error
        const row = worksheet.getRow(output.row); //  ensure row exists
        const cell = row.getCell(output.col);     //  get cell from the row
        cell.value = replaceText; //change the value of that cell (the value we assign to the replaceText assign to cell.value)
        await workbook.xlsx.writeFile(pathFile); // but by executing this line we will save the changes the excel file by using the writeFile function
    } else {
        console.log(`"${searchText}" not found in the sheet.`); //  added message when value is not found
    }
}

async function readExcel(worksheet, searchText) {
     //Currently, no valid row or column is selected or found.... with the help of this we can access within a block
    let output = { row: -1, col: -1 };
    worksheet.eachRow((row, rowNumber) => //Goes through each row in the Excel sheet.  For each row, this function runs. rowNumber gives row index (like 1, 2, 3...)
    {
        row.eachCell((cell, colNumber) => //Now, go through each cell in the current row.  For each cell in the row, this function runs.
        {
            if (cell.value === searchText) { // it gives the index of the apple in the excel table 
                console.log(output.row = rowNumber); //console.log(cell.rowNumber); both are same 
                console.log(output.col = colNumber); //console.log(cell.columnNumber); both are work same
            }
            // console.log(cell.value); // Print the value inside the current cell (like "Rajpal", "test@gmail.com", etc.)
        });

    });
    return output; //it returns the output of that function and then declear this value to that readExcel function 
}

writeExcelTest("Iphone", "Apple", "C:/Users/Pankaj/Downloads/ExceldownloadTest.xlsx"); // calling that function----> just by enter the parameters value  ("text we search","text we replace","file path of that file")
// to run the excel file or code write----------->node <file name.js> like----> node excelDemo.js and it prints the value 
