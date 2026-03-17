/**
 * @file Delete_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiName.Delete
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to delete a defined name from a worksheet.
 * It adds a defined name to a range, then retrieves and deletes it, and finally displays a confirmation message.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как удалить определенное имя из листа.
 * Он добавляет определенное имя к диапазону, затем извлекает и удаляет его, а затем отображает сообщение с подтверждением.
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
        // This example deletes the DefName object.
        
        // How to remove custom DefName from a worksheet.
        
        // Delete previously added DefName. 
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        Api.AddDefName("numbers", "Sheet1!$A$1:$B$1");
        let defName = Api.GetDefName("numbers");
        defName.Delete();
        worksheet.GetRange("A3").SetValue("The name 'numbers' of the range A1:B1 was deleted.");
        
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
