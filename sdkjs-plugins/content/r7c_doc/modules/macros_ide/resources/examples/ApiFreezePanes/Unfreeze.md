/**
 * @file Unfreeze_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiFreezePanes.Unfreeze
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to unfreeze all panes in the worksheet.
 * It freezes the first column, then unfreezes all panes, and then displays the location of the frozen panes (which should be empty).
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как разморозить все области на листе.
 * Он закрепляет первый столбец, затем размораживает все области, а затем отображает расположение закрепленных областей (которое должно быть пустым).
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
        // This example freezes first column then unfreeze all panes in the worksheet.
        
        // How to unfreeze columns from freezed panes.
        
        // Add freezed panes then unfreeze the first column and show all freezed ones' location to prove it.
        
        Api.SetFreezePanesType('column');
        let worksheet = Api.GetActiveSheet();
        let freezePanes = worksheet.GetFreezePanes();
        freezePanes.Unfreeze();
        let range = freezePanes.GetLocation();
        worksheet.GetRange("A1").SetValue("Location: ");
        worksheet.GetRange("B1").SetValue(range + "");
        
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
