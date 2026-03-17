/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.FISHERINV
 * 
 *  Демонстрация использования метода FISHERINV класса ApiWorksheetFunction
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
        // This example shows how to get an inverse of the Fisher transformation: if y = FISHER(x), then FISHERINV(y) = x.
        
        // How to get an inverse of the Fisher transformation.
        
        // Use function to find out an inverse of Fisher transformation.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let ans = func.FISHERINV(0.56);
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
