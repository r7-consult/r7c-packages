/**
 * @file ReplaceCurrentImage_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.ReplaceCurrentImage
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to replace the current image with a new one.
 * It adds an image to the worksheet and then replaces it with a different image from a URL.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как заменить текущее изображение новым.
 * Он добавляет изображение на лист, а затем заменяет его другим изображением из URL-адреса.
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
        // This example replaces the image with a new one.
        
        // How to replace an image to another one.
        
        // Replace an image from one to another using their urls.
        
        let worksheet = Api.GetActiveSheet();
        let drawing = worksheet.AddImage("https://api.R7 Office.com/content/img/docbuilder/examples/coordinate_aspects.png", 60 * 36000, 35 * 36000, 0, 2 * 36000, 0, 3 * 36000);// todo_example we don't have method ApiDrawing.Select() which is necessary for this example
        worksheet.ReplaceCurrentImage("https://helpcenter.R7 Office.com/images/Help/GettingStarted/Documents/big/EditDocument.png", 60 * 36000, 35 * 36000);
        
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
