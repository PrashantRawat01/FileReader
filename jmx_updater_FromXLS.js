let xlsx = require('xlsx');
let fs = require('fs');
var testJMX_FilePath = 'TestPlan.jmx'

function jmxUpdate() {
    try {
        var workbook = xlsx.readFile('./ThreadController.xls'); //Reading the excel data file
        var range = xlsx.utils.decode_range(workbook.Sheets[workbook.SheetNames[0]]['!ref']); //Recipients
        var newValue = fs.readFileSync(testJMX_FilePath, "utf8");

        for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
            var ChangeNodeName = workbook.Sheets[workbook.SheetNames[0]][xlsx.utils.encode_cell({ r: rowNum, c: 0 })].v;
            var ChangeNodeValue = workbook.Sheets[workbook.SheetNames[0]][xlsx.utils.encode_cell({ r: rowNum, c: 1 })].v;

            switch (ChangeNodeName) {
                case 'SBS_GetAccountServiceStartupTime':
                    newValue = newValue.replace(/SBS_GetAccountServiceStartupTime/, ChangeNodeValue);
                    break;
                case 'SBS_GetAccountServiceThreadCount':
                    newValue = newValue.replace(/SBS_GetAccountServiceThreadCount/, ChangeNodeValue);
                    break;
                case 'SBS_GetAccountServiceThreadRampUpTime':
                    newValue = newValue.replace(/SBS_GetAccountServiceThreadRampUpTime/, ChangeNodeValue);
                    break;
                case 'SBS_GetUnitHoldingStartupTime':
                    newValue = newValue.replace(/SBS_GetUnitHoldingStartupTime/, ChangeNodeValue);
                    break;
                case 'SBS_GetUnitHoldingThreadCount':
                    newValue = newValue.replace(/SBS_GetUnitHoldingThreadCount/, ChangeNodeValue);
                    break;
                case 'SBS_GetUnitHoldingThreadRampUpTime':
                    newValue = newValue.replace(/SBS_GetUnitHoldingThreadRampUpTime/, ChangeNodeValue);
                    break;
                case 'SBS_GetClientHoldingThreadCount':
                    newValue = newValue.replace(/SBS_GetClientHoldingThreadCount/, ChangeNodeValue);
                    break;
                case 'SBS_GetClientThreadCount':
                    newValue = newValue.replace(/SBS_GetClientThreadCount/, ChangeNodeValue);
                    break;
                case 'SBS_GetClientThreadRampUpTime':
                    newValue = newValue.replace(/SBS_GetClientThreadRampUpTime/, ChangeNodeValue);
                    break;
                default:
                    console.log('There is no tag found for replacement: ' + ChangeNodeName);
            }
        }
        new fs.writeFileSync(testJMX_FilePath, newValue);
        console.log('Done!')

    } catch (err) {
        console.log("Error in updating the Jmx file due to error: " + err);
    }
}

jmxUpdate()