/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.ECMA_CEILING
 * 
 *  Демонстрация использования метода ECMA_CEILING класса ApiWorksheetFunction
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
        // This example shows how to round the number up to the nearest multiple of significance. Negative numbers are rounded towards zero.
        
        // How to round up the number.
        
        // Use function to round up a number to the nearest multiple of significance.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.ECMA_CEILING(1.567, 0.1));
        
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
