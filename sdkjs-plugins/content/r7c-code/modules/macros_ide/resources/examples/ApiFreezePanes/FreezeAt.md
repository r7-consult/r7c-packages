/**
 * @file FreezeAt_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiFreezePanes.FreezeAt
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to freeze a specified range in the top-and-left-most pane of the worksheet.
 * It gets the freeze panes object, defines a range, and then freezes the panes at that range.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как закрепить указанный диапазон в самой верхней и левой области листа.
 * Он получает объект закрепленных областей, определяет диапазон, а затем закрепляет области в этом диапазоне.
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
        // This example freezes the specified range in top-and-left-most pane of the worksheet.
        
        // How to freeze a specified range of panes.
        
        // Get freeze panes and freeze the specified part.
        
        let worksheet = Api.GetActiveSheet();
        let freezePanes = worksheet.GetFreezePanes();
        let range = Api.GetRange('H2:K4');
        freezePanes.FreezeAt(range);
        
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
