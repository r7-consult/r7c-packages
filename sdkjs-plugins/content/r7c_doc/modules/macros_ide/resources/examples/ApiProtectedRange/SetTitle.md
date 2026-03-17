/**
 * @file SetTitle_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiProtectedRange.SetTitle
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the title for a protected range.
 * It adds a protected range to the worksheet and then changes its title.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить заголовок для защищенного диапазона.
 * Он добавляет защищенный диапазон на лист, а затем изменяет его заголовок.
 *
 * @returns {void}
 *
 * @see https://r7-consult.ru/
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
        // This example changes the the user protected range title.
        
        // How to set a title for a protected range.
        
        // Rename a title of a protected range.
        
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
