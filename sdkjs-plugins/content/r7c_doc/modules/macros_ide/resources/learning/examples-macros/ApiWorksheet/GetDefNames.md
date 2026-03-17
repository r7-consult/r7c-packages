/**
 * @file GetDefNames_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetDefNames
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get an array of ApiName objects.
 * It adds two defined names to the worksheet and then displays their names.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить массив объектов ApiName.
 * Он добавляет два определенных имени на лист, а затем отображает их имена.
 *
 * @returns {void}
 *
 * @see https://r7-consult.com/
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
        // This example shows how to get an array of ApiName objects.
        
        // How to get all def names.
        
        // Get all def names as an array.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        worksheet.GetRange("A2").SetValue("A");
        worksheet.GetRange("B2").SetValue("B");
        worksheet.AddDefName("numbers", "Sheet1!$A$1:$B$1");
        worksheet.AddDefName("letters", "Sheet1!$A$2:$B$2");
        let defNames = worksheet.GetDefNames();
        worksheet.GetRange("A4").SetValue("DefNames: " + defNames[0].GetName() + ", " + defNames[1].GetName());
        
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
