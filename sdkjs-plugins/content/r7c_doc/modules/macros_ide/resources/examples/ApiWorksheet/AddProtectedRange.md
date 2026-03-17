/**
 * @file AddProtectedRange_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.AddProtectedRange
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to add a new protected range to the worksheet.
 * It sets values in a range and then adds a protected range to it.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как добавить новый защищенный диапазон на лист.
 * Он устанавливает значения в диапазоне, а затем добавляет к нему защищенный диапазон.
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
        // This example adds a new protected range.
        
        // How to add the protected ApiRange object.
        
        // Insert a protected range to the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        worksheet.AddProtectedRange("protectedRange", "Sheet1!$A$1:$B$1");
        
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
