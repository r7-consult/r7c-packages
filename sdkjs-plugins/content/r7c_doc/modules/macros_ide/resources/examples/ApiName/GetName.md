/**
 * @file GetName_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiName.GetName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the name of a defined name.
 * It adds a defined name to a range, then retrieves it and displays its name.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить имя определенного имени.
 * Он добавляет определенное имя к диапазону, затем извлекает его и отображает его имя.
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
        // This example shows how to get a type of the ApiName class.
        
        // How to retrieve name of ApiName class object.
        
        // Get name of a specified object.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        Api.AddDefName("numbers", "Sheet1!$A$1:$B$1");
        let defName = Api.GetDefName("numbers");
        worksheet.GetRange("A3").SetValue("Name: " + defName.GetName());
        
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
