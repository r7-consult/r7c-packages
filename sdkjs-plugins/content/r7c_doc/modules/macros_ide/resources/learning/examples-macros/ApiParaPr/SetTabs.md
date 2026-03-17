/**
 * @file SetTabs_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParaPr.SetTabs
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set a sequence of custom tab stops for a paragraph.
 * It creates a shape, gets its content, sets custom tab stops with different positions and alignments,
 * and then adds text with tabs to demonstrate their effect.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить последовательность пользовательских табуляций для абзаца.
 * Он создает фигуру, получает ее содержимое, устанавливает пользовательские табуляции с различными позициями и выравниваниями,
 * а затем добавляет текст с табуляциями для демонстрации их эффекта.
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
        // This example sets a sequence of custom tab stops which will be used for any tab characters in the paragraph.
        
        // How to change sizes of tabs between paragraphs.
        
        // Customize all kind of tabs indicating sizes.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 150 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let paraPr = paragraph.GetParaPr();
        paraPr.SetTabs([1440, 2880, 4320], ["left", "center", "right"]);
        paragraph.AddTabStop();
        paragraph.AddText("Custom tab - 1 inch left");
        paragraph.AddLineBreak();
        paragraph.AddTabStop();
        paragraph.AddTabStop();
        paragraph.AddText("Custom tab - 2 inches center");
        paragraph.AddLineBreak();
        paragraph.AddTabStop();
        paragraph.AddTabStop();
        paragraph.AddTabStop();
        paragraph.AddText("Custom tab - 3 inches right");
        
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
