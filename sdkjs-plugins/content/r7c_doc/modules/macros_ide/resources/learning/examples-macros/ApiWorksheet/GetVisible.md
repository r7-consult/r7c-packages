/**
 * @file GetVisible_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetVisible
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the state of sheet visibility.
 * It sets the active worksheet to visible and then displays its visibility status.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить состояние видимости листа.
 * Он устанавливает активный лист видимым, а затем отображает его статус видимости.
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
        // This example shows how to get the state of sheet visibility.
        
        // How to get visibility of the worksheet.
        
        // Find out whether a sheet is visible or not and display it in the sheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetVisible(true);
        let isVisible = worksheet.GetVisible();
        worksheet.GetRange("A1").SetValue("Visible: ");
        worksheet.GetRange("B1").SetValue(isVisible);
        
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
