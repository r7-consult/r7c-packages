/**
 * @file SetIndFirstLine_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParagraph.SetIndFirstLine
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the first line indentation of a paragraph.
 * It creates a shape, adds text to the first paragraph, sets its first line indentation, and then adds another paragraph without indentation.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить отступ первой строки абзаца.
 * Он создает фигуру, добавляет текст в первый абзац, устанавливает отступ первой строки, а затем добавляет еще один абзац без отступа.
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
        // This example sets the paragraph first line indentation.
        
        // How to change first sentence indentation of the paragraph.
        
        // Get a paragraph from the shape's content then add a text specifying the first line indentation.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        paragraph.AddText("This is a paragraph with the indent of 1 inch set to the first line. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes.");
        paragraph.SetIndFirstLine(1440);
        paragraph = Api.CreateParagraph();
        paragraph.AddText("This is a paragraph without any indent set to the first line. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes.");
        content.Push(paragraph);
        
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
