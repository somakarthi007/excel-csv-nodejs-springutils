require("dotenv").config({"path": "./.env"})
const xlsx = require('node-xlsx'); 
const express = require("express");
const app = express();
const port = process.env.PORT || 5000;
const fs = require('fs');
const {ExcelDateToJSDate} = require("./utils")

app.use(express.json({limit: '50mb'}))

app.post("/xlsx/:sheetNumber", (req, res) => {

    const {encodedStr} = req.body;
    const n = encodedStr.length
    let buff = new Buffer.alloc(4*(n/3), encodedStr, 'base64');
    fs.writeFileSync('final.xlsx', buff);
    console.log("File written successfully!")

    var obj = xlsx.parse('final.xlsx'); // parses a file
    var rows = [];
    var writeStr = "";
    
    //convert the respective sheet to csv
    const sheetNumber = req.params.sheetNumber;
    console.log("Sheet Number: " + sheetNumber);
    
    if(sheetNumber != null){
        var sheet = obj[sheetNumber];
        //loop through all rows in the sheet
        for(var j = 0; j < sheet['data'].length; j++){
            //   cols = []
            //     for(var k = 0; k < sheet['data'][j].length; k++){
            //         //add the col to the cols array
            //         let col_data = sheet['data'][j][k];
            //         console.log("ColData: "  + col_data)
            //         if(col_data != null && col_data != undefined && col_data.toString().split('.').length > 1 && col_data.toString().split('.')[1].length > 10){
            //             cols.push(ExcelDateToJSDate(parseFloat(col_data)))
            //         }
            //         cols.push(sheet['data'][j][k])
            //     }
            //     //add the row to the rows array
            //     rows.push(cols);
        rows.push(sheet['data'][j]);
        }
    }

    //creates the csv string to write it to a file
    for(var i = 0; i < rows.length; i++)
    {
        writeStr += rows[i].join(",") + "\n";
    }

    console.log(process.memoryUsage());

    if (fs.existsSync("final.xlsx")) {
        fs.unlinkSync("final.xlsx")
    }
    console.log("ok!");
    res.send(writeStr);
   
})

app.post("/xlsb/:sheetNumber", (req, res) => {

    const {encodedStr} = req.body;
    const n = encodedStr.length
    let buff = new Buffer.alloc(4*(n/3), encodedStr, 'base64');
    fs.writeFileSync('final.xlsb', buff);
    console.log("File written successfully!")

    var obj = xlsx.parse('final.xlsb'); // parses a file
    var rows = [];
    var writeStr = "";
    
    //convert the respective sheet to csv
    const sheetNumber = req.params.sheetNumber;
    console.log("Sheet Number: " + sheetNumber);
    
    if(sheetNumber != null){
        var sheet = obj[sheetNumber];
        //loop through all rows in the sheet
        for(var j = 0; j < sheet['data'].length; j++)
        {
                //add the row to the rows array
                cols = []
                for(var k = 0; k < sheet['data'][j].length; k++){
                    //add the col to the cols array
                    let col_data = sheet['data'][j][k];
                    if(col_data != null && col_data != undefined && col_data.toString().split('.').length > 1 && col_data.toString().split('.')[1].length > 9){
                        cols.push(ExcelDateToJSDate(parseFloat(col_data)))
                    }
                    cols.push(sheet['data'][j][k])
                }
                //add the row to the rows array
                rows.push(cols);
                rows.push(sheet['data'][j]);
        }
    }

    //creates the csv string to write it to a file
    for(var i = 0; i < rows.length; i++)
    {
        writeStr += rows[i].join(",") + "\n";
    }

    console.log(process.memoryUsage());

    if (fs.existsSync("final.xlsb")) {
        fs.unlinkSync("final.xlsb")
    }

    console.log("ok!");
    res.send(writeStr);
   
})

app.listen(port, () => console.log(`Listening to port ${port}`))