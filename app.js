// Require the library, create a instance of a Workbook class and one worksheet
var xl = require('excel4node');
const workbook = require('excel4node/distribution/lib/workbook');

var wb = new xl.Workbook();

var ws = wb.addWorksheet('Sheet 1');

// Define the column headings
var headingColumnNames = [
    'Name',
    'Capital',
    'Area',
    'Currencies'
]

// Define all the styles of the worksheet 
var styleHeader = wb.createStyle({
    alignment: {
        horizontal : 'center'
    },
    font: {
        color: '#4F4F4F',
        size: 16,
        bold: true
    }
})

var styleColumnHeadings = wb.createStyle({
    font: {
        color: '#808080',
        size: 12,
        bold: true
    }
})

// Make the first header
ws.cell(1,1,1,4,true).string('Countries List').style(styleHeader)

// Make all the column headings
let headingColumnIndex = 1;
headingColumnNames.forEach(element => {
    ws.cell(2, headingColumnIndex++).string(element).style(styleColumnHeadings)
})


// Create the xlsx file
wb.write('test.xlsx')