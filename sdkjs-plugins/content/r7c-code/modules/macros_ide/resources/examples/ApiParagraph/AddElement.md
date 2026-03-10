/**
 * @file AddElement_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParagraph.AddElement
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to add a run to a paragraph within a shape.
 * It creates a shape, gets its content, retrieves the first paragraph, creates a new run with text,
 * and then adds this run to the paragraph.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как добавить запуск в абзац внутри фигуры.
 * Он создает фигуру, получает ее содержимое, извлекает первый абзац, создает новый запуск с текстом,
 * а затем добавляет этот запуск в абзац.
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
        // This example adds a Run to the paragraph.
        
        // How to add text to the paragraph.
        
        // Get the paragraph from the shape and change its text.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        run.AddText("This is just a sample text run. Nothing special.");
        paragraph.AddElement(run);
        
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
