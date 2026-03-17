/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.PRICEDISC
 * 
 *  Демонстрация использования метода PRICEDISC класса ApiWorksheetFunction
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
        // This example shows how to return the price per $100 face value for a discounted security.
        
        // How to return the price per $100 face value for a discounted security.
        
        // Use a function to return the price per $100 face value for a discounted security.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.PRICEDISC("1/1/2017", "6/1/2019", 0.05, 105, 1));
        
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
