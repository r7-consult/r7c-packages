/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.MUNIT
 * 
 *  Демонстрация использования метода MUNIT класса ApiWorksheetFunction
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
        // This example shows how to return the unit matrix for the specified dimension.
        
        // How to get a unit matrix.
        
        // Use a function to get a unit matrix specifying dimension.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.MUNIT(3));
        
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
