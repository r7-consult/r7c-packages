/**
 * @file GetAddress_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetAddress
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the address of a range.
 * It sets values in cells A1 and B1, gets the address of cell A1, and then displays the address.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить адрес диапазона.
 * Он устанавливает значения в ячейках A1 и B1, получает адрес ячейки A1, а затем отображает адрес.
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
        // This example shows how to get the range address.
        
        // How to get an address of a range.
        
        // Get an address of one range and set it for another one.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        let address = worksheet.GetRange("A1").GetAddress(true, true, "xlA1", false);
        worksheet.GetRange("A3").SetValue("Address: ");
        worksheet.GetRange("B3").SetValue(address);
        
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
