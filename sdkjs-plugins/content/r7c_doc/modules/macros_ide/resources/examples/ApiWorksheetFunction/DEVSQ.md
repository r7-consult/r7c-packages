/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.DEVSQ
 * 
 *  Демонстрация использования метода DEVSQ класса ApiWorksheetFunction
 * https://r7-consult.ru/
 */

(function() {
    'use strict';
    
    try {
        // Initialize R7 Office API
        const api = Api;
        if (!api) {
            throw new Error('R7 Office API not available');
        }
        
        // Original code enhanced with error handling:
        // This example shows how to get the sum of squares of deviations of data points from their sample mean.
        
        // How to get sum of squares of deviations.
        
        // Use function to get the sum of squares of deviations of data points from their sample mean.
        
        let worksheet = Api.GetActiveSheet();
        let argumentsArrA = [34, 244];
        let argumentsArrB = [769, 445];
        let argumentsArrC = [76, 677];
        let argumentsArrD = [89, 56];
        let argumentsArrE = [98, 13];
        
        // Place the numbers in cells
        for (let a = 0; a < argumentsArrA.length; a++) {
            worksheet.GetRange("A" + (a + 1)).SetValue(argumentsArrA[a]);
        }
        for (let b = 0; b < argumentsArrB.length; b++) {
            worksheet.GetRange("B" + (b + 1)).SetValue(argumentsArrB[b]);
        }
        for (let c = 0; c < argumentsArrC.length; c++) {
            worksheet.GetRange("C" + (c + 1)).SetValue(argumentsArrC[c]);
        }
        for (let d = 0; d < argumentsArrD.length; d++) {
            worksheet.GetRange("D" + (d + 1)).SetValue(argumentsArrD[d]);
        }
        for (let e = 0; e < argumentsArrE.length; e++) {
            worksheet.GetRange("E" + (e + 1)).SetValue(argumentsArrE[e]);
        }
        
        // Analyze the range of data 
        let func = Api.GetWorksheetFunction();
        let ans = func.DEVSQ(worksheet.GetRange("A1:E2"));
        worksheet.GetRange("E3").SetValue(ans);
        
        // Success notification
        console.log('Macro executed successfully');
        
    } catch (error) {
        console.error('Macro execution failed:', error.message);
        // Optional: Show error to user
        if (typeof Api !== 'undefined' && Api.GetActiveSheet) {
            const sheet = Api.GetActiveSheet();
            if (sheet) {
                sheet.GetRange('A1').SetValue('Error: ' + error.message);
            }
        }
    }
})();
