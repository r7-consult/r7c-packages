/**
 * @file GetAllImages_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetAllImages
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get all images from the sheet.
 * It adds an image to the worksheet, and then retrieves all images to display the class type of the first image.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить все изображения с листа.
 * Он добавляет изображение на лист, а затем извлекает все изображения для отображения типа класса первого изображения.
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
        // This example shows how to get all images from the sheet.
        
        // How to get all images.
        
        // Get all images as an array.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.AddImage("https://api.R7 Office.com/content/img/docbuilder/examples/coordinate_aspects.png", 60 * 36000, 35 * 36000, 0, 2 * 36000, 0, 3 * 36000);
        let images = worksheet.GetAllImages();
        let classType = images[0].GetClassType();
        worksheet.GetRange("A10").SetValue("Class Type = " + classType);
        
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
