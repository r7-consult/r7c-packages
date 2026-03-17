/**
 * @file GetRefersToRange_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiName.GetRefersToRange
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the ApiRange object that a defined name refers to.
 * It adds a defined name to a range, then retrieves the range object and makes its content bold.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект ApiRange, на который ссылается определенное имя.
 * Он добавляет определенное имя к диапазону, затем извлекает объект диапазона и делает его содержимое жирным.
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
        // This example shows how to get the ApiRange object by its name.
        
        // How to get a range knowig its defname.
        
        // Find a range by its name and change its properties.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        Api.AddDefName("numbers", "$A$1:$B$1");
        let defName = Api.GetDefName("numbers");
        let range = defName.GetRefersToRange();
        range.SetBold(true);
        
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
