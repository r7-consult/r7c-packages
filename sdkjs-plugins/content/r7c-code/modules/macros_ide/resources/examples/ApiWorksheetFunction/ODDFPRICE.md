/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.ODDFPRICE
 * 
 *  Демонстрация использования метода ODDFPRICE класса ApiWorksheetFunction
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
        // This example shows how to return the price per $100 face value of a security with an odd first period.
        
        // How to get the price per $100 face value of a security with an odd first period.
        
        // Use a function to return the price per $100 face value of a security.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.ODDFPRICE("1/1/2017", "6/1/2019", "12/1/2016", "3/15/2017", 0.05, 0.09, 100, 2));
        
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
