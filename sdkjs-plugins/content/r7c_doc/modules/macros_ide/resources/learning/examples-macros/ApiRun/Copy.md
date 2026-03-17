/**
 * @file Copy_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRun.Copy
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to create a copy of a run.
 * It creates a shape, adds a run with text, copies it, and then adds the copy to the paragraph.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как создать копию запуска.
 * Он создает фигуру, добавляет запуск с текстом, копирует его, а затем добавляет копию в абзац.
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
        // This example reates a copy of the run.
        
        // How to create a text run object and its copy.
        
        // Create an ApiRun and its copy and add it into paragraph.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        run.AddText("This is just a sample text that was copied. ");
        paragraph.AddElement(run);
        let copyRun = run.Copy();
        paragraph.AddElement(copyRun);
        
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
