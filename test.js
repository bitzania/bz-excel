
var bz_excel = require('./index.js');
var path = require("path");

var data = [
    {field1: 1, field2:"100.00", field3: true, field4: new Date()},
    {field1: 2, field2:"11.00", field3: false, field4: new Date()},
    {field1: 3, field2:"12.00", field3: true, field4: new Date()},
    {field1: 4, field2:"13.00", field3: true, field4: new Date()},
    {field1: 5, field2:"14.00", field3: null, field4: new Date()},
    {field1: 6, field2:"15.00", field3: true, field4: new Date()},
]

console.log("WRITING EXCEL FILE : ", path.join(__dirname, "./xl.xlsx"));
bz_excel.createExcel(data, ["field1", "field3", "field2:number", "field4"], path.join(__dirname, "./withField.xlsx"))
.then((data)=> {
    console.log(data);
})
.catch((err)=> {
    console.log(err);
})
// bz_excel.createExcel(data, null, path.join(__dirname, "./default.xlsx"));