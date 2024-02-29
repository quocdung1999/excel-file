const puppet = require('puppeteer')
var xlsx = require('xlsx')
var fs = require('fs')
const {setTimeout} = require("node:timers/promises")
const {PDFDocument, StandardFonts, rgb} = require('pdf-lib')
const f = async function () {
    const browser = await puppet.launch({headless: true, slowMo: 200})
    const page = await browser.newPage()
    const dataPathExcelToRead = "Sample.xlsx";
    const wb = xlsx.readFile(dataPathExcelToRead);
    const sheetName = wb.SheetNames[0]
    const sheetValue = wb.Sheets[sheetName]

    const htmlData = xlsx.utils.sheet_to_html(sheetValue)
    fs.writeFile('excelToHtml.html', htmlData, err => {
        console.log("data is successfully converted")
    })
    await setTimeout(1000)
    await page.goto(`file://${process.cwd()}/excelToHtml.html`, {"waitUntil":"networkidle2"})
    console.log(process.cwd())
    await page.pdf({path:"./ExcelToPDF.pdf", format: "A4", printBackground: true})
    await browser.close()
}

 f()