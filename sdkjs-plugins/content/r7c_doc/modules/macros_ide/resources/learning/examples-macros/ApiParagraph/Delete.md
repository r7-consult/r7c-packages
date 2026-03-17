/**
 * @file Delete_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParagraph.Delete
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to delete a paragraph from a shape's content.
 * It creates a shape, adds a paragraph with text, and then deletes the paragraph.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как удалить абзац из содержимого фигуры.
 * Он создает фигуру, добавляет абзац с текстом, а затем удаляет абзац.
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
        // This example deletes the paragraph.
        
        // How to remove a paragraph.
        
        // Delete the paragraph from the shape content.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        content.RemoveAllElements();
        let paragraph = Api.CreateParagraph();
        paragraph.AddText("This is just a sample text.");
        content.Push(paragraph);
        paragraph.Delete();
        worksheet.GetRange("A9").SetValue("The paragraph from the shape content was removed.");
        
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
