/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.ROWS
 * 
 *  Демонстрация использования метода ROWS класса ApiWorksheetFunction
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
        // This example shows how to return the number of rows in a range.
        
        // How to count number of rows.
        
        // Use a function to count number of rows.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let column1 = [13, 14, 15];
        let column2 = [23, 24, 25];
        
        for (let i = 0; i < column1.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(column1[i]);
        }
        for (let j = 0; j < column2.length; j++) {
            worksheet.GetRange("B" + (j + 1)).SetValue(column2[j]);
        }
        
        let range = worksheet.GetRange("A1:B3");
        worksheet.GetRange("C3").SetValue(func.ROWS(range));
        
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
