/**
 * @file SetApplicationId_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiOleObject.SetApplicationId
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the application ID for an OLE object.
 * It adds an OLE object to the worksheet and then sets its application ID.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить идентификатор приложения для объекта OLE.
 * Он добавляет объект OLE на лист, а затем устанавливает идентификатор его приложения.
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
        // This example sets the application ID to the current OLE object.
        
        // How to set application id of OLE object.
        
        // Add Ole object, set its application id and display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let oleObject = worksheet.AddOleObject("https://api.R7 Office.com/content/img/docbuilder/examples/ole-object-image.png", 130 * 36000, 90 * 36000, "https://youtu.be/SKGz4pmnpgY", "asc.{38E022EA-AD92-45FC-B22B-49DF39746DB4}", 0, 2 * 36000, 4, 3 * 36000);
        oleObject.SetApplicationId("asc.{E5773A43-F9B3-4E81-81D9-CE0A132470E7}");
        
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
