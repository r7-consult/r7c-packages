/**
 * @file SetDoubleStrikeout_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiTextPr.SetDoubleStrikeout
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to specify that the contents of the run are displayed with two horizontal lines through each character displayed on the line.
 * It creates a shape, adds a run with text, sets its double strikethrough property, and then adds the run to the paragraph.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как указать, что содержимое запуска отображается двумя горизонтальными линиями через каждый символ, отображаемый в строке.
 * Он создает фигуру, добавляет запуск с текстом, устанавливает его свойство двойного зачеркивания, а затем добавляет запуск в абзац.
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
        // This example specifies that the contents of the run are displayed with two horizontal lines through each character displayed on the line.
        
        // How to strike out a text with two lines.
        
        // Cross out a text with two lines.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        let textProps = run.GetTextPr();
        textProps.SetFontSize(30);
        textProps.SetDoubleStrikeout(true);
        paragraph.SetJc("left");
        run.AddText("This is a sample text inside the shape struck out with two lines using the text properties.");
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
