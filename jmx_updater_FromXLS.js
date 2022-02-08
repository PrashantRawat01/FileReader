//Created by: Prashant Rawat
//Created Date: 06/02/2022
//purpose: To replace content of any file.!

let xlsx = require('xlsx');
let xlsxs = require('xlsx');
let fs = require('fs');
var testJMX_FilePath = 'TestPlan.jmx'

function jmxUpdate() {
    try {
        var configWorkbook = xlsx.readFile('./JMeterConfigs.xls'); //Reading the excel data file
        var threadWorkbook = xlsx.readFile('./ThreadController.xls'); //Reading the excel data file
        var ThreadFileSheet ;
        //var ThreadFileRange = xlsxs.utils.decode_range(threadWorkbook.Sheets[threadWorkbook.SheetNames[0]]['!ref']);
        
        var newValue = fs.readFileSync(testJMX_FilePath, "utf8");
        var configFileRange = xlsx.utils.decode_range(configWorkbook.Sheets[configWorkbook.SheetNames[0]]['!ref']); //Recipients
        for (let rowNum = configFileRange.s.r; rowNum <= configFileRange.e.r; rowNum++) {
            var ChangeNodeName = configWorkbook.Sheets[configWorkbook.SheetNames[0]][xlsx.utils.encode_cell({ r: rowNum, c: 0 })].v;
            var ChangeNodeValue = configWorkbook.Sheets[configWorkbook.SheetNames[0]][xlsx.utils.encode_cell({ r: rowNum, c: 1 })].v;

            switch (ChangeNodeName) {
                case 'PackToRun':
                    if (ChangeNodeValue == 'RampUp') { ThreadFileSheet = 0};
                    if (ChangeNodeValue == 'Spike') { ThreadFileSheet = 1};
                    if (ChangeNodeValue == 'Soak') { ThreadFileSheet = 2};
                    break;
                case 'JMeterScriptLocation':
                    newValue = newValue.replace(/JMeterScriptLocation/, ChangeNodeValue);
                    break;
                case 'ReportLocation':
                    newValue = newValue.replace(/ReportLocation/, ChangeNodeValue);
                    break;
                case 'RunDate&Time':
                    newValue = newValue.replace(/RunDate&Time/, ChangeNodeValue);
                    break;
                default:
                    console.log('There is no tag found to replace from Config Sheet: ' + ChangeNodeName);
            }
        }
        var ThreadFileRange = xlsx.utils.decode_range(threadWorkbook.Sheets[threadWorkbook.SheetNames[ThreadFileSheet]]['!ref'])
        for (let rowNum = ThreadFileRange.s.r; rowNum <= ThreadFileRange.e.r; rowNum++) {
            ChangeNodeName = threadWorkbook.Sheets[threadWorkbook.SheetNames[ThreadFileSheet]][xlsx.utils.encode_cell({ r: rowNum, c: 0 })].v;
            ChangeNodeValue = threadWorkbook.Sheets[threadWorkbook.SheetNames[ThreadFileSheet]][xlsx.utils.encode_cell({ r: rowNum, c: 1 })].v;

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
                    console.log('There is no tag found for replacement from ThreadController: ' + ChangeNodeName);
            }
        }
        new fs.writeFileSync(testJMX_FilePath, newValue);
        console.log('Done!')

    } catch (err) {
        new fs.writeFileSync(testJMX_FilePath, newValue);
        console.log("Error in updating the Jmx file due to error: " + err);
    }
}

jmxUpdate()
