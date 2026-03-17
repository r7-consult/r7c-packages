/**
 * @file SetData_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiOleObject.SetData
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the data for an OLE object.
 * It adds an OLE object to the worksheet and then sets its data to a new URL.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить данные для объекта OLE.
 * Он добавляет объект OLE на лист, а затем устанавливает его данные на новый URL.
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
        // This example sets the data to the current OLE object.
        
        // How to change content of OLE object.
        
        // Add Ole object, set its data and display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let oleObject = worksheet.AddOleObject("https://api.R7 Office.com/content/img/docbuilder/examples/ole-object-image.png", 130 * 36000, 90 * 36000, "https://youtu.be/SKGz4pmnpgY", "asc.{38E022EA-AD92-45FC-B22B-49DF39746DB4}", 0, 2 * 36000, 4, 3 * 36000);
        oleObject.SetData("https://youtu.be/eJxpkjQG6Ew");
        
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
