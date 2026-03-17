/**
 * @file GetProtectedRange_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetProtectedRange
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get an object that represents the protected range.
 * It adds a protected range to the worksheet, and then retrieves it to set its title.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект, представляющий защищенный диапазон.
 * Он добавляет защищенный диапазон на лист, а затем извлекает его, чтобы установить его заголовок.
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
        // This example shows how to get an object that represents the protected range.
        
        // How to get protected range.
        
        // Get protected range and set its title.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.AddProtectedRange("protectedRange", "Sheet1!$A$1:$B$1");
        let protectedRange = worksheet.GetProtectedRange("protectedRange");
        protectedRange.SetTitle("protectedRangeNew");
        
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
