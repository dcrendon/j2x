// require packages
const flatten = require('flat');
const XLSX = require("xlsx");
const prompt = require('prompt-sync')({sigint: true});
const fs = require('fs');

// set var to track number of records
let num = 0

//create array for flat data
var flatdata = [];

// ask user for file name
console.log("Make sure your JSON file is in the same directory as this exe");
const filename = prompt('Enter the name of the file that needs to be converted: ');
const path = `./${filename}`

// check if file provided exists
if (!fs.existsSync(path)) {
    console.log('\x1b[31m%s\x1b[0m', 'The file you entered was not found.')
    prompt('Press [Enter] to close the application: ')
    return
}

// get json data from file
var filedata = fs.readFileSync(path);
filedata = JSON.parse(filedata);

// loop through raw json, flatten each object, push to new array, convert to json string
filedata.forEach(user => {
    num++
    let flatuser = flatten(user,{
        delimiter: '-'
    })
    flatdata.push(flatuser)
});

flatdata = JSON.stringify(flatdata)
 
var flatjson = JSON.parse(flatdata);

// convert to xls
const worksheet = XLSX.utils.json_to_sheet(flatjson);
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, "data");

// write xls file
XLSX.writeFile(workbook, "data.xlsx");

// output results
console.log('\x1b[32m%s\x1b[0m', `${num} records processed`)
console.log('\x1b[32m%s\x1b[0m', `Conversion completed`)

prompt('Press [ENTER] to close the application: ');