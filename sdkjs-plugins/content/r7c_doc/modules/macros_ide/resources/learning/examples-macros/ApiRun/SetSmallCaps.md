/**
 * @file SetSmallCaps_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRun.SetSmallCaps
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to specify that all the small letter characters in this text run are formatted for display only as their capital letter character equivalents which are two points smaller than the actual font size specified for this text.
 * It creates a shape, adds a run with normal text, and then adds another run with small capitalized text.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как указать, что все строчные символы в этом текстовом запуске форматируются для отображения только как их эквиваленты заглавных букв, которые на два пункта меньше фактического размера шрифта, указанного для этого текста.
 * Он создает фигуру, добавляет запуск с обычным текстом, а затем добавляет еще один запуск с текстом, написанным заглавными буквами.
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
        // This example specifies that all the small letter characters in this text run are formatted for display only as their capital letter character equivalents which are two points smaller than the actual font size specified for this text.
        
        // How to make text characters uncapitalized.
        
        // Create a text run object, update its style by making its letters uncapitalized.
        
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
        run.SetSmallCaps(true);
        run.AddText("This is a text run with the font set to small capitalized letters.");
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
