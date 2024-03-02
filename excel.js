const ExcelJS = require("exceljs");
const {PDFDocument, PDFFont, StandardFonts, StandardFontValues, rgb} = require('pdf-lib')
const {writeFileSync} = require("node:fs");

// Padding in pixel
const paddingH = 50
const paddingW = 50

const defaultH = 16
const HtoPixel = 1
const defaultW = 10
const WtoPixel = 6.5

async function readExcelStyles(excelFilePath) {
    try {
        const pdfDoc = await PDFDocument.create()

        const timesRoman = await pdfDoc.embedFont(StandardFonts.TimesRoman)
        const timesRomanBold = await pdfDoc.embedFont(StandardFonts.TimesRoman)
        const timesRomanItalic = await pdfDoc.embedFont(StandardFonts.TimesRoman)
        const timesRomanBoldItalic = await pdfDoc.embedFont(StandardFonts.TimesRoman)

        const page = pdfDoc.addPage([550,750])
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);

        workbook.eachSheet(async (worksheet, sheetId) => {
            console.log(`Sheet Name: ${worksheet.name}`);

            for (const image of worksheet.getImages()) {
                // console.log('processing image row', image.range.tl.nativeRow, 'col', image.range.tl.nativeCol, 'imageId', image.imageId);
                // // fetch the media item with the data (it seems the imageId matches up with m.index?)
                const img = workbook.model.media.find(m => m.index === image.imageId);
                console.log("OKKK")
                let pngImage = await pdfDoc.embedPng(img.buffer)
                console.log(pngImage.width, pngImage.height)
                page.drawImage(pngImage, {x: 20, y: 20, width: pngImage.width, height: pngImage.height})
                //writeFileSync(`${image.range.tl.row}.${image.range.tl.col}.${image.range.br.row}.${image.range.br.col}.${img.name}.${img.extension}`, img.buffer);
            }
            let currX = 0
            let currY = 740
            // worksheet.columns.forEach((column, colNumber) => {
            //     console.log(`Column ${colNumber} Width: ${column.width}`);
            // });
            //  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            //      console.log(`Row ${rowNumber} Width: ${row.height}`);
            //  });
            // worksheet.eachRow({includeEmpty: true}, (row, rowNumber) => {
            //     row.eachCell({includeEmpty: true}, (cell, colNumber) => {
            //
            //         console.log(`  Cell R${rowNumber}C${colNumber}:`);
            //         console.log(`    Value: ${JSON.stringify(cell.value)}`);
            //         // if (cell.value !== null) {
            //         //     currX += 30
            //         //     //currY -= 10
            //         //     page.drawText(cell.value, {x: currX, y: currY, size: cell.style.font.size, font: timesRoman})
            //         //
            //         // }
            //
            //
            //         console.log(`    Style:`, cell.style);
            //     });
            // });

            page.drawRectangle({x: 0, y: 730, width: 40, height: 20, color: rgb(233 / 255, 1, 0)})


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
