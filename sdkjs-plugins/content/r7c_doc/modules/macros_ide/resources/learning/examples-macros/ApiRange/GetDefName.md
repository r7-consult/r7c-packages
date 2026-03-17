/**
 * @file GetDefName_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetDefName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the ApiName object of a range.
 * It sets values in a range, adds a defined name to it, and then displays the name of the defined name.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект ApiName диапазона.
 * Он устанавливает значения в диапазоне, добавляет к нему определенное имя, а затем отображает имя определенного имени.
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
        // This example shows how to get the ApiName object of the range.
        
        // How to find out a range name.
        
        // Get a range, get its name and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        Api.AddDefName("numbers", "Sheet1!$A$1:$B$1");
        let range = worksheet.GetRange("A1:B1");
        let defName = range.GetDefName();
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
