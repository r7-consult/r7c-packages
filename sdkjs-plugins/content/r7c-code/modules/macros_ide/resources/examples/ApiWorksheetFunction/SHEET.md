/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.SHEET
 * 
 *  Демонстрация использования метода SHEET класса ApiWorksheetFunction
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
        // This example shows how to return the sheet number of the reference sheet.
        
        // How to get a sheet number.
        
        // Use a function to get a sheet index.
        
        const worksheet = Api.GetActiveSheet(); 
        let func = Api.GetWorksheetFunction();
        let result = func.SHEET("Sheet1");
        worksheet.GetRange("C3").SetValue(result);
        
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
