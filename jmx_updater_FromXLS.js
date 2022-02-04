var XLSX = require('xlsx');
const fs = require('fs');
var testJMX_FilePath = 'TestPlan.jmx'

async function jmxUpdate() {

    try{
        var workbook = XLSX.readFile('./ThreadController.xls'); //Reading the excel data file
        var range = XLSX.utils.decode_range(workbook.Sheets[workbook.SheetNames[0]]['!ref']);         //Recipients
       // var secondCell = workbook.Sheets[workbook.SheetNames[0]][XLSX.utils.encode_cell({ r: 1, c: 1 })];
        //var firstCell;
        var ChangeNodeName
        var ChangeNodeValue
        var serviceName

        for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
            
            ChangeNodeName = workbook.Sheets[workbook.SheetNames[0]][XLSX.utils.encode_cell({ r: rowNum, c: 0 })];
            //console.log(ChangeNodeName.v);
            if (ChangeNodeName.v == 'ThreadName') {
                serviceName = workbook.Sheets[workbook.SheetNames[0]][XLSX.utils.encode_cell({ r: rowNum, c: 1 })];
                //console.log(serviceName.v);
                continue;
            }
           ChangeNodeValue = workbook.Sheets[workbook.SheetNames[0]][XLSX.utils.encode_cell({ r: rowNum, c: 1 })];
           //console.log(ChangeNodeValue.v);
           var nodeValueToBeUpdate = serviceName.v+ChangeNodeName.v;
           console.log(nodeValueToBeUpdate);


           fs.readFile(testJMX_FilePath, 'utf-8', function(err, data) {
            if (err) throw err;
            var newValue = data.replace('/'+ nodeValueToBeUpdate +'/gim', ChangeNodeValue.v);
        
            fs.writeFile(testJMX_FilePath, newValue, 'utf-8', function(err, data) {
                if (err) throw err;
                console.log('Done!');
            })
            //sleep(1);
            })
        

        }
    }catch(err){
        console.log("Error in updating the Jmx file due to error: " + err);
    }
    }


    // (async function () {
    //     const fileContent = await fs.readFile(testJMX_FilePath);
    //     const records = parse(fileContent, {columns: true});
    //     console.log(records)
    // })();


async function sleep(msec) {
        console.log("Sleeping for Miliseconds " + msec); 
        return new Promise(resolve => setTimeout(resolve, msec));
    }



    jmxUpdate()