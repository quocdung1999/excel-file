const ExcelJS = require("exceljs");
const {writeFileSync} = require("fs");

async function readExcelStyles(excelFilePath) {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);

        workbook.eachSheet((worksheet, sheetId) => {
            console.log(`Sheet Name: ${worksheet.name}`);

            for (const img of workbook.model.media) {
                // console.log('processing image row', image.range.tl.nativeRow, 'col', image.range.tl.nativeCol, 'imageId', image.imageId);
                // // fetch the media item with the data (it seems the imageId matches up with m.index?)
                // const img = workbook.model.media.find(m => m.index === image.imageId);
                console.log("OKKK")
                writeFileSync(` ${img.name}.${img.extension}`, img.buffer);
            }

            // worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            //     row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            //         console.log(`  Cell R${rowNumber}C${colNumber}:`);
            //         console.log(`    Value: ${JSON.stringify(cell.value)}`);
            //         console.log(`    Style:`, cell.style);
            //     });
            // });
        });
    } catch (error) {
        console.error('Error reading Excel file:', error.message);
    }
}

// Example usage
const excelFilePath = './Book1.xlsx';
readExcelStyles(excelFilePath);
