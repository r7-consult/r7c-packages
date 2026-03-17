/**
 * @file GetWorksheet_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetWorksheet
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the Worksheet object that represents the worksheet containing the specified range.
 * It sets a value in a range, gets its worksheet, and then displays the worksheet's name.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект Worksheet, представляющий лист, содержащий указанный диапазон.
 * Он устанавливает значение в диапазоне, получает его лист, а затем отображает имя листа.
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
        // This example shows how to get the Worksheet object that represents the worksheet containing the specified range.
        
        // How to get a worksheet where a range is contained in.
        
        // Get a worksheet from its range and show its name.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1:C1");
        range.SetValue("1");
        let oSheet = range.GetWorksheet();
        worksheet.GetRange("A3").SetValue("Worksheet name: " + oSheet.GetName());
        
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
