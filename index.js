// Require library
var xl = require('excel4node');

module.exports = {
    

    createExcel: function(rows, fields, filename) {
        return new Promise((resolve, reject)=> {
            // Create a new instance of a Workbook class
            var wb = new xl.Workbook();

            // Add Worksheets to the workbook
            var ws = wb.addWorksheet('Sheet 1');

            var types = [];
            if (!fields) {
                fields = Object.keys(rows[0]);
            }


            // write fields
            for (var j = 0; j< fields.length; j++) {
                var field = fields[j];

                console.log("FIELD", field);
                if (field.indexOf(':')>-1) {
                    var parts = field.split(':');
                    field = parts[0].toString().trim();
                    types.push(parts[1].toString().trim());
                }
                else types.push("auto");
                ws.cell(1, j+1)["string"](field);
                fields[j] = field;
            }

            for (var i =0; i< rows.length; i ++) {
                var row = rows[i];
                for (var j = 0; j< fields.length; j++) {
                    var field = fields[j];
                    var otype = types[j];
                    var value = row[field] || '';
                    // value = value.toString();
                    var type = typeof value;

                    if (otype != "auto") type = otype;
                    switch (type) {
                        case "boolean":
                            type = "bool";
                        break;

                        case "number": 
                            value = value * 1;
                        break;

                        case "object":
                            if (Object.prototype.toString.call(value) === '[object Date]') {
                                type = "date";
                                // console.log("ISDATE");
                            }
                            else {
                                type = "string";
                                value = value.toString();
                            }
                        break;

                    }
                    console.log("type", type, "value", value, Object.prototype.toString.call(value)=== '[object Date]');
                    ws.cell(i+2, j+1)[type](value);
                }

            }

            wb.write(filename, (err, stats) => {
                if (err) {
                    reject(err);
                } else {
                    resolve(filename);
                }
            });
        }) 
        
    }


}