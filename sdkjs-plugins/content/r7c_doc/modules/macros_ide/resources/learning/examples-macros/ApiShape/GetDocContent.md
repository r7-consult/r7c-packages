/**
 * @file GetDocContent_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiShape.GetDocContent
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the shape inner contents where a paragraph or text runs can be inserted.
 * It creates a shape, gets its document content, removes all existing elements, creates a new paragraph with text, and then adds this paragraph to the shape's content.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить внутреннее содержимое фигуры, куда можно вставлять абзацы или текстовые фрагменты.
 * Он создает фигуру, получает ее содержимое документа, удаляет все существующие элементы, создает новый абзац с текстом, а затем добавляет этот абзац в содержимое фигуры.
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
        // This example shows how to get the shape inner contents where a paragraph or text runs can be inserted.
        
        // How to get content of ApiShape.
        
        // Get content of ApiShape, remove all its elements and add a new paragraph to it.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetDocContent();
        content.RemoveAllElements();
        let paragraph = Api.CreateParagraph();
        paragraph.SetJc("left");
        paragraph.AddText("We removed all elements from the shape and added a new paragraph inside it.");
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
