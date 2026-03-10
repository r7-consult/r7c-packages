/**
 * @file AddElement_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiDocumentContent.AddElement
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to add a paragraph to the document content within a shape.
 * It creates a shape, removes all existing elements from its content, creates a new paragraph with text,
 * and then adds this paragraph to the shape's content.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как добавить абзац в содержимое документа внутри фигуры.
 * Он создает фигуру, удаляет все существующие элементы из ее содержимого, создает новый абзац с текстом,
 * а затем добавляет этот абзац в содержимое фигуры.
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
        // This example adds a paragraph in document content.
        
        // How to add text to the document using ApiParagraph.
        
        // Update the document content adding a paragraph to it.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        content.RemoveAllElements();
        let paragraph = Api.CreateParagraph();
        paragraph.AddText("We removed all elements from the shape and added a new paragraph inside it.");
        content.AddElement(paragraph);
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
