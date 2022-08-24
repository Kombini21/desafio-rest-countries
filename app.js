// Require the node-fetch (for the API call), excel4node and the readline module (for the user input), create a instance of a Workbook class and one worksheet
var fetch = require('node-fetch');
var xl = require('excel4node');
const readline = require('readline');

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

// Create the first header and apply the style
ws.cell(1,1,1,4,true).string('Countries List').style(styleHeader)

// Create all the column headings and apply the style
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

        // Fetch the json data and if it's undefined, define it to -
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
            countryArea = countryArea.toLocaleString("en-US")
            countryArea = countryArea.replaceAll(',', '.') + ',00'
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



// Define the function to fetch the data and create the final xlsx file 

async function fetchAndCreateXlsx(array, filename) {

    //For each country in the array, fetch the data and write in the worksheet
    for (i = 0; i < array.length; i++) {

        let data = await countryApiFetchData(array[i]);
        writeFetchedData(data, i + 3)
    }

    // Create the xlsx file
    wb.write(`./output/${filename}.xlsx`)
    console.log('File created!')
}

// Define a function to create a xlsx file with all the countries from the API

async function fetchAndCreateAllXlsx() {

    await fetch(`https://restcountries.com/v2/all`)
    .then(response => {
        return response.json()
    })
    .then(body => {

        let allCountriesArr = [];
        for (i = 0; i < body.length; i++) {
            let palavra = body[i].name.replace(/\([^()]*\)$/g, '')
            allCountriesArr.push(palavra)
            
        }
        return allCountriesArr
    }).then(arr => {
        console.log('Making file!')
        fetchAndCreateXlsx(arr, 'allCountries')
    })
}





// Define the readline for user input, and start the application


const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

// Create the terminal response based on the user input

rl.question('Hello, welcome to the country data fetcher!\n Choose one of the execution methods below! - \n 1 - default -- create a default xlsx file \n 2 - all countries -- return a xlsx file with data of all countries \n 3 - user input -- input your the countries you want data of \n 4 -- exit - exit app\n\n', choices => {

    if (choices == '1') {
            console.log('loading')
            let countriesArr = ['Afghanistan', 'Åland Islands','Albania',' Algeria', 'American Samoa', 'Andorra', 'Angola', 'Anguilla', 'Antarctica', 'Antigua and Barbuda', 'Argentina', 'Armenia', 'Aruba', 'Australia', 'Azerbaijan', 'Bahamas', 'Bahrain', 'Bangladesh', 'Barbados', 'Belarus']
            fetchAndCreateXlsx(countriesArr, 'countriesDefault')
    }

    else if (choices == '2') {
        console.log('loading');
        fetchAndCreateAllXlsx();
    }

    else if (choices == '3'){

        console.log('teste')
    }


    else if (choices == '4') {
        console.log('Closing app!')
    }

    rl.close();
});



