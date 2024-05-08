//Load Excel file
var XLSX_file = require("xlsx");
var workbook = XLSX_file.readFile("Release Due.xlsx");

//Get the first worksheet
let worksheet = workbook.Sheets[workbook.SheetNames[0]];

//console.log(worksheet);

const firstColumn = 'F';
const endColumn = 'W';

//This is to create a json of all of the worksheet per row
//Ex. raw_data[0] is row 1 (the column titles)
const raw_data = XLSX_file.utils.sheet_to_json(worksheet, {header: 1});
//console.log(raw_data.length);

// Get the range of rows in the worksheet
const range = XLSX_file.utils.decode_range(worksheet['!ref']);
const rowCount = range.e.r - range.s.r + 1; // Number of rows
console.log(rowCount);

//Create objects in arrays
//Start the loop on column 5 which is 'End of Inspection'
//End the loop on column 14 which is 'Base unit of measure'
//raw_data[y] is Per row
//raw_data[y][x] is per item in the specific row

let y = 0;
for (let x = 0;x<100;x++){
    
    
    
    for (let x = 0;x<raw_data[y].length;x++){
        if (x>=5&&x<13&&x!=10){
            console.log('x is '+x+'raw_data['+y+']'+'['+x+']'+raw_data[y][x]);
        }
        //console.log('x is '+x+'raw_data['+y+']'+'['+x+']'+raw_data[y][x]);
            
    }
    
    y++;

    
}





