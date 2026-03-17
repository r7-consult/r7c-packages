/**
 * @file GetSpacingLineRule_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParagraph.GetSpacingLineRule
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the paragraph line spacing rule.
 * It creates a shape, adds text to the first paragraph, sets its line spacing, and then displays the spacing rule.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить правило межстрочного интервала абзаца.
 * Он создает фигуру, добавляет текст в первый абзац, устанавливает его межстрочный интервал, а затем отображает правило интервала.
 *
 * @returns {void}
 *
 * @see https://r7-consult.ru/
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
        // This example shows how to get the paragraph line spacing rule.
        
        // How to get the spacing information of the paragraph lines.
        
        // Create a paragraph, set the spacing line between the sentences and show it. 
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 80 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        paragraph.SetSpacingLine(3 * 240, "auto");
        paragraph.AddText("Paragraph 1. Spacing: 3 times of a common paragraph line spacing.");
        paragraph.AddLineBreak();
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.AddLineBreak();
        let spacingRule = paragraph.GetSpacingLineRule();
        paragraph.AddText("Spacing line rule: " + spacingRule);
        
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
