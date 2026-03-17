/**
 * @file GetAllOleObjects_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetAllOleObjects
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get all OLE objects from the sheet.
 * It adds an OLE object to the worksheet, and then retrieves all OLE objects to display the application ID of the first one.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить все объекты OLE с листа.
 * Он добавляет объект OLE на лист, а затем извлекает все объекты OLE для отображения идентификатора приложения первого из них.
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
        // This example shows how to get all OLE objects from the sheet.
        
        // How to get all OLE objects images.
        
        // Get all OLE objects as an array.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.AddOleObject("https://i.ytimg.com/vi_webp/SKGz4pmnpgY/sddefault.webp", 130 * 36000, 90 * 36000, "https://youtu.be/SKGz4pmnpgY", "asc.{38E022EA-AD92-45FC-B22B-49DF39746DB4}", 0, 2 * 36000, 4, 3 * 36000);
        let oleObjects = worksheet.GetAllOleObjects();
        let appId = oleObjects[0].GetApplicationId();
        worksheet.GetRange("A1").SetValue("The application ID for the current OLE object: " + appId);
        
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
