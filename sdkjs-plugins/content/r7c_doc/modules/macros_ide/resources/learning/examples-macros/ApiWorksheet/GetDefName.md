/**
 * @file GetDefName_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetDefName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the ApiName object by the worksheet name.
 * It adds a defined name to a range, and then retrieves and displays its name.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект ApiName по имени листа.
 * Он добавляет определенное имя к диапазону, а затем извлекает и отображает его имя.
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
        // This example shows how to get the ApiName object by the worksheet name.
        
        // How to get def name object.
        
        // Get ApiName object using its name.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        worksheet.AddDefName("numbers", "Sheet1!$A$1:$B$1");
        let defName = worksheet.GetDefName("numbers");
        worksheet.GetRange("A3").SetValue("DefName: " + defName.GetName());
        
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
