/**
 * @file AddTabStop_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRun.AddTabStop
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to add a tab stop to a run.
 * It creates a shape, adds a run with text, adds three tab stops, and then adds more text after the tab stops.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как добавить табуляцию к запуску.
 * Он создает фигуру, добавляет запуск с текстом, добавляет три табуляции, а затем добавляет больше текста после табуляции.
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
        // This example adds a tab stop to the run.
        
        // How to add a tab to a sentence.
        
        // Break two lines of a text run with a tab. 
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        run.SetFontSize(30);
        run.AddText("This is just a sample text. After it three tab stops will be added.");
        run.AddTabStop();
        run.AddTabStop();
        run.AddTabStop();
        run.AddText("This is the text which starts after the tab stops.");
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
