/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.VARA
 * 
 *  Демонстрация использования метода VARA класса ApiWorksheetFunction
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
        // This example shows how to estimate variance based on a sample, including logical values and text. Text and the false logical value have the value 0; the true logical value has the value 1.
        
        // How to estimate variance based on a sample considering logical and text data types.
        
        // Use a function to estimate variance not ignoring logical and text values.
        
        
        let worksheet = Api.GetActiveSheet();
        
        // Data array
        let data = [
            [1, 0, 0, false],
            ["text", 1, 0, 0],
            [2, 3, true, 7],
            [6, 8, 10, 12]
        ];
        
        // Place the data in cells A1:D4
        for (let i = 0; i < data.length; i++) {
            for (let j = 0; j < data[i].length; j++) {
                worksheet.GetRangeByNumber(i, j).SetValue(data[i][j]);
            }
        }
        
        // Calculate the letA of the range A1:D4 and place the result in cell E6
        let func = Api.GetWorksheetFunction();
        let letaResult = func.VARA(worksheet.GetRange("A1:D4"));
        worksheet.GetRange("E6").SetValue(letaResult);
        
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
