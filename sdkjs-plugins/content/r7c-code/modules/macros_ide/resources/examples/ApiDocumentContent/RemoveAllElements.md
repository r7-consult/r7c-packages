/**
 * @file RemoveAllElements_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiDocumentContent.RemoveAllElements
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to remove all elements from the document content within a shape.
 * It creates a shape, adds a sample paragraph, then removes all elements, and finally adds a new paragraph.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как удалить все элементы из содержимого документа внутри фигуры.
 * Он создает фигуру, добавляет образец абзаца, затем удаляет все элементы и, наконец, добавляет новый абзац.
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
        // This example removes all the elements from the current document or from the current document content.
        
        // How to clear a document.
        
        // Delete all elements from a document.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        paragraph.AddText("This is just a sample paragraph.");
        content.RemoveAllElements();
        paragraph = Api.CreateParagraph();
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
