/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.VDB
 * 
 *  Демонстрация использования метода VDB класса ApiWorksheetFunction
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
        // This example shows how to return a depreciation of an asset for any specified period, including partial periods, using the double-declining balance method or some other method specified.
        
        // How to estimate depreciation of an asset for any specified period.
        
        // Use a depreciation of an asset for any specified period including partial periods.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.VDB(3500, 500, 5, 1, 3, 2, false));
        
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
