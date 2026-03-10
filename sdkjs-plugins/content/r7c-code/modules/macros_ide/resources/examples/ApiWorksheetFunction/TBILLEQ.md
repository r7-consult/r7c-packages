/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.TBILLEQ
 * 
 *  Демонстрация использования метода TBILLEQ класса ApiWorksheetFunction
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
        // This example shows how to return the bond-equivalent yield for a treasury bill.
        
        // How to return the bond-equivalent yield for a treasury bill.
        
        // Use a function to calculate the bond-equivalent yield.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.TBILLEQ("1/1/2018", "11/20/2018", "8.00%"));
        
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
