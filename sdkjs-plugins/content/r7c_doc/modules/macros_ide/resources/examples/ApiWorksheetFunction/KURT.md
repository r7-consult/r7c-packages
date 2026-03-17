/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.KURT
 * 
 *  Демонстрация использования метода KURT класса ApiWorksheetFunction
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
        // This example shows how to return the kurtosis of a data set.
        
        // How to know a data set kurtosis.
        
        // Use a function to find out kurtosis of a data set.
        
        const worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let kurtosis = func.KURT(3, 89, 34, 2, 45, 4, 45, 13);
        worksheet.GetRange("C1").SetValue(kurtosis);
        
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
