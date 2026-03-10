/**
 * @file Delete_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRun.Delete
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to delete a run.
 * It creates a shape, adds a run with text to the first paragraph, and then deletes the run.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как удалить запуск.
 * Он создает фигуру, добавляет запуск с текстом в первый абзац, а затем удаляет запуск.
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
        // This example deletes the run.
        
        // How to delete a text run object.
        
        // Create the ApiRun object, add it into the paragraph and remove it from it.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        run.AddText("This is just a sample text.");
        paragraph.AddElement(run);
        run.Delete();
        worksheet.GetRange("A9").SetValue("The run from the shape content was removed.");
        
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
