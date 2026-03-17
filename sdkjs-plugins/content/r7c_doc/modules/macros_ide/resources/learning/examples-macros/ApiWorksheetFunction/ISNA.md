/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.ISNA
 * 
 *  Демонстрация использования метода ISNA класса ApiWorksheetFunction
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
        // This example shows how to check whether a value is N/A, and returns true or false. 
        
        // How to check if the cell contains N/A value.
        
        // Use a function to check whether a range data is an N/A value.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.ISNA("#N/A"));
        worksheet.GetRange("A2").SetValue(func.ISNA(255));
        worksheet.GetRange("A3").SetValue(func.ISNA("www.example.com"));
        
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
