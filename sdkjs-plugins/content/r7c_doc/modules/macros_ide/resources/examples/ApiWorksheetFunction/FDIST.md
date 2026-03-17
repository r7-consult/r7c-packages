/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.FDIST
 * 
 *  Демонстрация использования метода FDIST класса ApiWorksheetFunction
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
        // This example shows how to return the (right-tailed) F probability distribution (degree of diversity) for two data sets.
        
        // How to get an F probability distribution of two datasets.
        
        // Use function to get F probability distribution (degree of diversity).
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let ans = func.FDIST(10, 6, 4);
        worksheet.GetRange("B2").SetValue(ans);
        
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
