const ExcelJS = require("exceljs");
const {PDFDocument} = require('pdf-lib')
const {writeFileSync} = require("node:fs");

async function readExcelStyles(excelFilePath) {
    try {
        const pdfDoc = await PDFDocument.create()
        const page = pdfDoc.addPage([550,750])
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);

        workbook.eachSheet((worksheet, sheetId) => {
            console.log(`Sheet Name: ${worksheet.name}`);

            // for (const image of worksheet.getImages()) {
            //     // console.log('processing image row', image.range.tl.nativeRow, 'col', image.range.tl.nativeCol, 'imageId', image.imageId);
            //     // // fetch the media item with the data (it seems the imageId matches up with m.index?)
            //      const img = workbook.model.media.find(m => m.index === image.imageId);
            //     console.log("OKKK")
            //     writeFileSync(`${image.range.tl.row}.${image.range.tl.col}.${image.range.br.row}.${image.range.br.col}.${img.name}.${img.extension}`, img.buffer);
            // }

            worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                    console.log(`  Cell R${rowNumber}C${colNumber}:`);
                    console.log(`    Value: ${JSON.stringify(cell.value)}`);
                    if (cell.value !== null) {
                        page.drawText(cell.value, {x : 10, y: 10, size: 10})
                    }

                    console.log(`    Style:`, cell.style);
                });
            });
        });

        const pdfBytes = await pdfDoc.save()
        writeFileSync("abc.pdf", pdfBytes)


    } catch (error) {
        console.error('Error reading Excel file:', error.message);
    }




}

// Example usage
const excelFilePath = './Book1.xlsx';
readExcelStyles(excelFilePath);
