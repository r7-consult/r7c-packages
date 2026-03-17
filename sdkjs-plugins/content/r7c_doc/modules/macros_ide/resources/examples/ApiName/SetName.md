/**
 * @file SetName_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiName.SetName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the name of a defined name.
 * It adds a defined name to a range, renames it, and then displays the new name.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить имя определенного имени.
 * Он добавляет определенное имя к диапазону, переименовывает его, а затем отображает новое имя.
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
        // This example sets a string value representing the object name.
        
        // How to rename an object.
        
        // Set a new name for an object and display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        Api.AddDefName("name", "Sheet1!$A$1:$B$1");
        let defName = Api.GetDefName("name");
        defName.SetName("new_name");
        let newDefName = Api.GetDefName("new_name");
        worksheet.GetRange("A3").SetValue("The new name of the range: " + newDefName.GetName());
        
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
