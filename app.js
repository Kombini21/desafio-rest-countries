// Require the node-fetch (for the API call) and the excel4node library, create a instance of a Workbook class and one worksheet
var fetch = require('node-fetch');
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

// Make the first header and apply the style
ws.cell(1,1,1,4,true).string('Countries List').style(styleHeader)

// Make all the column headings and apply the style
let headingColumnIndex = 1;
headingColumnNames.forEach(element => {
    ws.cell(2, headingColumnIndex++).string(element).style(styleColumnHeadings)
})



// Define the API request function 

async function countryApiFetchData(country) {

    //Fetch the country data
    let response = await fetch(`https://restcountries.com/v3.1/name/${country}`)
    .then(response => {
        return response.json()
    })
    .then(body => {

        let countryName
        let countryCapital
        let countryArea
        let countryCurrency 


        if (typeof body[0].name.common === 'undefined') {
            countryName = '-'
        }else {countryName = body[0].name.common}


        if (typeof body[0].capital === 'undefined') {
            countryCapital = '-'
        }else {countryCapital = body[0].capital[0]}


        if (typeof body[0].area === 'undefined') {
            countryArea = '-'
        }else {
            countryArea = body[0].area
            countryArea = countryArea.toString()
        }

        if (typeof body[0].currencies === 'undefined') {
            countryCurrency = '-'
        }else {
            countryCurrency = body[0].currencies
            countryCurrency = Object.keys(countryCurrency)[0]
        }
        
        //Save the data in a array and return it
        let countryData = []
        countryData.push(countryName, countryCapital, countryArea, countryCurrency)
        return countryData
    })

    return response 
}



// Define the function to write the fetched data on the workshee

async function writeFetchedData(dataArr, rowIndex) {
    let ColumnIndex = 1;
    dataArr.forEach(element => {
        ws.cell(rowIndex, ColumnIndex++).string(element)
    })
}






async function makeXlsxFile(array, filename) {

}




// Create the xlsx file
//wb.write('test.xlsx')




/////////////////////////////////////// TESTE
//countryApiFetchData('Antarctica').then(res => { console.log(res)})