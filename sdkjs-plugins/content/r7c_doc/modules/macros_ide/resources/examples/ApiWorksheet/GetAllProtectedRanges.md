/**
 * @file GetAllProtectedRanges_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetAllProtectedRanges
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get an object that represents all protected ranges.
 * It adds two protected ranges to the worksheet and then renames their titles.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект, представляющий все защищенные диапазоны.
 * Он добавляет два защищенных диапазона на лист, а затем переименовывает их заголовки.
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
        // This example shows how to get an object that represents all protected ranges.
        
        // How to get all protected ranges.
        
        // Get all protected ranges as an array.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.AddProtectedRange("protectedRange1", "Sheet1!$A$1:$B$1");
        worksheet.AddProtectedRange("protectedRange2", "Sheet1!$A$2:$B$2");
        let protectedRanges = worksheet.GetAllProtectedRanges();
        protectedRanges[0].SetTitle("protectedRangeNew1");
        protectedRanges[1].SetTitle("protectedRangeNew2");
        
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
