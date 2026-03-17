/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.DDB
 * 
 *  Демонстрация использования метода DDB класса ApiWorksheetFunction
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
        // This example shows how to the depreciation of an asset for a specified period using the double-declining balance method or some other method you specify. 
        
        // How to count the non-empty cells containing numbers in the field (column) of records in the database that match the conditions you specify.
        
        // Use function to count numbers from non-empty database records that met a condition specified.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.DDB(3500, 500, 5, 1, 2));
        
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
