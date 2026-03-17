/**
 * @file GetData_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiOleObject.GetData
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the string data from an OLE object.
 * It adds an OLE object to the worksheet and then displays its data.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить строковые данные из объекта OLE.
 * Он добавляет объект OLE на лист, а затем отображает его данные.
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
        // This example shows how to get the string data from the OLE object.
        
        // How to get ApiOleObject content as a string.
        
        // Get ApiOleObject data and display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let oleObject = worksheet.AddOleObject("https://api.R7 Office.com/content/img/docbuilder/examples/ole-object-image.png", 130 * 36000, 90 * 36000, "https://youtu.be/SKGz4pmnpgY", "asc.{38E022EA-AD92-45FC-B22B-49DF39746DB4}", 0, 2 * 36000, 4, 3 * 36000);
        let data = oleObject.GetData();
        worksheet.GetRange("A1").SetValue("The OLE object data: " + data);
        
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
