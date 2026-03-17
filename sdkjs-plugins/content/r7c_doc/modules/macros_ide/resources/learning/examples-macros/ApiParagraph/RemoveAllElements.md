/**
 * @file RemoveAllElements_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParagraph.RemoveAllElements
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to remove all elements from a paragraph.
 * It creates a shape, adds a run to the first paragraph, removes all elements from it, and then adds a new run.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как удалить все элементы из абзаца.
 * Он создает фигуру, добавляет запуск в первый абзац, удаляет из него все элементы, а затем добавляет новый запуск.
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
        // This example removes all the elements from the current paragraph.
        
        // How to clear a content from the paragraph.
        
        // Create a paragraph, add a text to it then delete all elements from it.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        run.AddText("This is the first text run in the current paragraph.");
        paragraph.AddElement(run);
        paragraph.RemoveAllElements();
        run = Api.CreateRun();
        run.AddText("We removed all the paragraph elements and added a new text run inside it.");
        paragraph.AddElement(run);
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
