/**
 * @file AddImage_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.AddImage
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to add an image to the sheet with the specified parameters.
 * It adds an image from a URL to the worksheet with a specified width, height, and position.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как добавить изображение на лист с указанными параметрами.
 * Он добавляет изображение из URL-адреса на лист с указанной шириной, высотой и положением.
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
        // This example adds an image to the sheet with the parameters specified.
        
        // How to add an image to the worksheet specifying its url and size.
        
        // Insert an image to the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.AddImage("https://api.R7 Office.com/content/img/docbuilder/examples/coordinate_aspects.png", 60 * 36000, 35 * 36000, 0, 2 * 36000, 0, 3 * 36000);
        
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
