/**
 * @file SetVertAlign_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRun.SetVertAlign
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to specify the alignment which will be applied to the contents of the current run in relation to the default appearance of the text run.
 * It creates a shape, adds a run with normal text, and then adds three more runs with text aligned to subscript, baseline, and superscript.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как указать выравнивание, которое будет применяться к содержимому текущего запуска относительно внешнего вида текстового запуска по умолчанию.
 * Он создает фигуру, добавляет запуск с обычным текстом, а затем добавляет еще три запуска с текстом, выровненным по подстрочному, базовому и надстрочному индексам.
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
        
        // How to set vertical alignment of a text object.
        
        // Create a text run object, specify its vertical alignment as "baseline", "subscript" or "superscript".
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        run.AddText("This is just a sample text. ");
        paragraph.AddElement(run);
        run = Api.CreateRun();
        run.SetVertAlign("subscript");
        run.AddText("This is a text run with the text aligned below the baseline vertically. ");
        paragraph.AddElement(run);
        run = Api.CreateRun();
        run.SetVertAlign("baseline");
        run.AddText("This is a text run with the text aligned by the baseline vertically. ");
        paragraph.AddElement(run);
        run = Api.CreateRun();
        run.SetVertAlign("superscript");
        run.AddText("This is a text run with the text aligned above the baseline vertically.");
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
