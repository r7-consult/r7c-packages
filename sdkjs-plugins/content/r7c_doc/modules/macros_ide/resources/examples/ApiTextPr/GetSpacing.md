/**
 * @file GetSpacing_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiTextPr.GetSpacing
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the text spacing property.
 * It creates a shape, adds a run with text, sets its spacing, and then displays the spacing value.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить свойство интервала между текстами.
 * Он создает фигуру, добавляет запуск с текстом, устанавливает его интервал, а затем отображает значение интервала.
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
        // This example gets a text spacing using its property.
        
        // How to find out space size of a text.
        
        // Get spacing size.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        run.AddText("This is just a sample text. ");
        run.AddText("The text properties are changed and the style is added to the paragraph. ");
        run.AddLineBreak();
        paragraph.AddElement(run);
        let textProps = run.GetTextPr();
        textProps.SetSpacing(80);
        paragraph = Api.CreateParagraph();
        let spacing = textProps.GetSpacing();
        paragraph.AddText("Text spacing: " + spacing);
        content.Push(paragraph);
        
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
