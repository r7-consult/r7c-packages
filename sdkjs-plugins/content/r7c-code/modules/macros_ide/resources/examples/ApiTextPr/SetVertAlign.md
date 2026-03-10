/**
 * @file SetVertAlign_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiTextPr.SetVertAlign
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to specify the alignment which will be applied to the contents of the current run in relation to the default appearance of the text run.
 * It creates a shape, adds a run with text, sets its vertical alignment to "superscript", and then adds the run to the paragraph.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как указать выравнивание, которое будет применяться к содержимому текущего запуска относительно внешнего вида текстового запуска по умолчанию.
 * Он создает фигуру, добавляет запуск с текстом, устанавливает его вертикальное выравнивание на «надстрочный индекс», а затем добавляет запуск в абзац.
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
        // This example specifies the alignment which will be applied to the contents of the current run in relation to the default appearance of the text run.
        
        // How to change vertical alignment of a text.
        
        // Make text superscript.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        let textProps = run.GetTextPr();
        textProps.SetVertAlign("superscript");
        paragraph.SetJc("left");
        run.AddText("This is a text inside the shape with vertical alignment set to 'superscript'.");
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
