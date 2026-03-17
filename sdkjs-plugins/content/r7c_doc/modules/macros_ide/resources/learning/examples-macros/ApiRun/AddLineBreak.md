/**
 * @file AddLineBreak_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRun.AddLineBreak
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to add a line break to a run position and start the next element from a new line.
 * It creates a shape, adds a run with text, adds a line break, and then adds more text on the new line.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как добавить разрыв строки в позицию выполнения и начать следующий элемент с новой строки.
 * Он создает фигуру, добавляет запуск с текстом, добавляет разрыв строки, а затем добавляет больше текста на новой строке.
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
        // This example adds a line break to the run position and starts the next element from a new line.
        
        // How to start a sentence on a new line.
        
        // Break two lines of a text run with a line. 
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        run.AddText("This is the text for the first line. Nothing special.");
        run.AddLineBreak();
        run.AddText("This is the text which starts from the beginning of the second line. ");
        run.AddText("It is written in two text runs, you need a space at the end of the first run sentence to separate them.");
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
