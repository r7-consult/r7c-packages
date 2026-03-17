/**
 * @file RemoveElement_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiDocumentContent.RemoveElement
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to remove an element from the document content within a shape by its position.
 * It creates a shape, adds multiple paragraphs, removes the third paragraph, and then adds a new paragraph to confirm the removal.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как удалить элемент из содержимого документа внутри фигуры по его позиции.
 * Он создает фигуру, добавляет несколько абзацев, удаляет третий абзац, а затем добавляет новый абзац для подтверждения удаления.
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
        // This example removes an element using the position specified.
        
        // How to remove an element from a document knowing its position in the document content.
        
        // Delete an element from a document and prove it by showing the difference.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        paragraph.AddText("This is paragraph #1.");
        for (let paraIncrease = 1; paraIncrease < 5; ++paraIncrease) {
            paragraph = Api.CreateParagraph();
            paragraph.AddText("This is paragraph #" + (paraIncrease + 1) + ".");
            content.Push(paragraph);
        }
        content.RemoveElement(2);
        paragraph = Api.CreateParagraph();
        paragraph.AddText("We removed paragraph #3, check that out above.");
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
