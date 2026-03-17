/**
 * @file SetRefersTo_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiName.SetRefersTo
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the formula that a defined name refers to.
 * It adds a defined name that refers to a range, sets the formula, and then displays a confirmation message.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить формулу, на которую ссылается определенное имя.
 * Он добавляет определенное имя, которое ссылается на диапазон, устанавливает формулу, а затем отображает сообщение с подтверждением.
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
        // This example sets a formula that the name is defined to refer to.
        
        // How to add a name and refer it to the formula.
        
        // Create new ranges and a formula, refer a name to the formula then display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        worksheet.GetRange("C1").SetValue("=SUM(A1:B1)");
        Api.AddDefName("summa", "Sheet1!$A$1:$B$1");
        let defName = Api.GetDefName("summa");
        defName.SetRefersTo("=SUM(A1:B1)");
        worksheet.GetRange("A3").SetValue("The name 'summa' refers to the formula from the cell C1.");
        
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
