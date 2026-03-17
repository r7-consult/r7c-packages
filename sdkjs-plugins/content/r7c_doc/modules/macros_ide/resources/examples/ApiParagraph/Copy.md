/**
 * @file Copy_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParagraph.Copy
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to create a copy of a paragraph.
 * It creates a shape, adds text to the first paragraph, copies it, and then adds the copy to the shape's content.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как создать копию абзаца.
 * Он создает фигуру, добавляет текст в первый абзац, копирует его, а затем добавляет копию в содержимое фигуры.
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
        // This example creates a paragraph copy.
        
        // How to create an identical paragraph.
        
        // Get a paragraph from the content of the shape create its copy and add it to the shape.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        paragraph.SetJc("left");
        paragraph.AddText("This is a text inside the shape aligned left.");
        paragraph.AddLineBreak();
        paragraph.AddText("This is a text after the line break.");
        let paragraph2 = paragraph.Copy();
        content.Push(paragraph2);
        
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
