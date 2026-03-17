/**
 * @file SetBullet_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParaPr.SetBullet
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set a bullet to a paragraph.
 * It creates a shape, gets its content, creates a bullet with a dash, and then sets the bullet to the paragraph properties.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить маркер для абзаца.
 * Он создает фигуру, получает ее содержимое, создает маркер с тире, а затем устанавливает маркер для свойств абзаца.
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
        // This example sets the bullet or numbering to the current paragraph.
        
        // How to add a dash bullet to the paragraph.
        
        // Numbering and adding custom bullet points to the text.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let paraPr = paragraph.GetParaPr();
        let bullet = Api.CreateBullet("-");
        paraPr.SetBullet(bullet);
        paragraph.AddText(" This is an example of the bulleted paragraph.");
        
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
