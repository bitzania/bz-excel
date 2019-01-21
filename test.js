
var bz_excel = require('./index.js');
var path = require("path");

var data = [
    {field1: 1, field2:"b", field3: true, field4: new Date()},
    {field1: 2, field2:"c", field3: false, field4: new Date()},
    {field1: 3, field2:"d", field3: true, field4: new Date()},
    {field1: 4, field2:"e", field3: true, field4: new Date()},
    {field1: 5, field2:"r", field3: null, field4: new Date()},
    {field1: 6, field2:"g", field3: true, field4: new Date()},
]

console.log("WRITING EXCEL FILE : ", path.join(__dirname, "./xl.xlsx"));
bz_excel.createExcel(data, ["field1", "field3", "field2", "field4"], path.join(__dirname, "./withField.xlsx"))
.then((data)=> {
    console.log(data);
})
.catch((err)=> {
    console.log(err);
})
bz_excel.createExcel(data, null, path.join(__dirname, "./default.xlsx"));