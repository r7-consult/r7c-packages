/**
 * @file GetLocation_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiFreezePanes.GetLocation
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the location of the frozen panes.
 * It freezes the first column, gets the location of the frozen panes, and then displays the address of the frozen range.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить расположение закрепленных областей.
 * Он закрепляет первый столбец, получает расположение закрепленных областей, а затем отображает адрес закрепленного диапазона.
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
        // This example freezes first column and get pastes a freezed range address into the table.
        
        // How to get location address of a freezed column.
        
        // Get an address of a column from freezed panes and display it in the worksheet.
        
        Api.SetFreezePanesType('column');
        let worksheet = Api.GetActiveSheet();
        let freezePanes = worksheet.GetFreezePanes();
        let range = freezePanes.GetLocation();
        worksheet.GetRange("A1").SetValue("Location: ");
        worksheet.GetRange("B1").SetValue(range.GetAddress());
        
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
