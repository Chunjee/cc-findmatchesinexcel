const
    fs = require('fs');
    
    
    // modules
var
    chalk = require('chalk'),
        // project specific
    XLSX = require('xlsx');


// global vars
var 
    amountCol =     "A",
    dateCol =       "C",
    projectCol =    "F",
    poCol =         "H",
    vendorCol =     "J",
    helperCol =     "I";

var matchHelperCounter = 0;

var workbook = XLSX.readFile('Data Sample for Nested match offsetting amounts.xlsx');
var ws = workbook.Sheets["Sheet1"]
var g_length = XLSX.utils.sheet_to_json(ws).length

console.log(chalk.green("starting search"))
for (let index = 0; index < g_length; index++) {
    var amountSearching = getCellValue(ws["A" + index])
    if (amountSearching < 0) {
        continue;
    }
    for (let index2 = 0; index2 < g_length; index2++) {
        var comparisonAmount = getCellValue(ws["A" + index2])
        // if the search and comparison value amount equal 0 sum, a potential match has been found
        if (amountSearching + comparisonAmount == 0) {
            var searchingString = "" + getCellValue(ws[projectCol + index]) + getCellValue(ws[poCol + index]) + getCellValue(ws[vendorCol + index])
            var comparisonString = "" + getCellValue(ws[projectCol + index2]) + getCellValue(ws[poCol + index2]) + getCellValue(ws[vendorCol + index2])
            if (searchingString.trim() == comparisonString.trim()) {
                // The vendor and project's match, now checking the dates
                searchingDate = getCellValue(ws[dateCol + index])
                comparisonDate = getCellValue(ws[dateCol + index2])
                if (compareDates(searchingDate,comparisonDate)) {
                    matchHelperCounter++
                    ws[helperCol + index].v = matchHelperCounter.toString()
                    ws[helperCol + index2].v = matchHelperCounter.toString()
                    console.log(chalk.yellow("match number " + matchHelperCounter + " was found at row " + index + " and " + index2));
                } else {
                    // console.log("offset found but the date was backwards")
                    // console.log("" + searchingDate + " - " + comparisonDate + "\n" + index + " - " + index2)
                }
            }
        }
    }
}
console.log(chalk.green("End of script reached"));
console.log(chalk.green(matchHelperCounter + " matching offsets found"));
// write file to disk
workbook.Sheets["Sheet2"] = ws
XLSX.writeFile(workbook, 'out.xlsb');
process.exit(1);



// Functions

function getCellValue(param_celladdress) {
    return (param_celladdress ? param_celladdress.v : undefined)
}

function compareDates(param_date1,param_date2) {
var date1 = Date.parse(param_date1);
var date2 = Date.parse(param_date2);
    if (date1 < date2) {
        return true;
    } else {
        return false;
    }
}
