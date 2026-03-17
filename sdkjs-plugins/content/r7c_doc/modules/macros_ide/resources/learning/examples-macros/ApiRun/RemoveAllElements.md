/**
 * @file RemoveAllElements_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRun.RemoveAllElements
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to remove all elements from a run.
 * It creates a shape, adds a run with text, removes all elements from it, and then adds new text to confirm the clearing.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как удалить все элементы из запуска.
 * Он создает фигуру, добавляет запуск с текстом, удаляет из него все элементы, а затем добавляет новый текст для подтверждения очистки.
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
        // This example removes all the elements from the run.
        
        // How to remove all text elements.
        
        // Create a text run object, add a text to it and clear its content.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        run.AddText("This is just a sample text.");
        run.RemoveAllElements();
        run.AddText("All elements from this run were removed before adding this text.");
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
