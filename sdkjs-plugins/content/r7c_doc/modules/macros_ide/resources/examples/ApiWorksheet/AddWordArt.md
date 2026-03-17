/**
 * @file AddWordArt_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.AddWordArt
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to add a Text Art object to the sheet with the specified parameters.
 * It creates text properties, a solid fill, and a stroke, and then adds a WordArt object to the worksheet with these properties.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как добавить объект Text Art на лист с указанными параметрами.
 * Он создает текстовые свойства, сплошную заливку и обводку, а затем добавляет объект WordArt на лист с этими свойствами.
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
        // This example adds a Text Art object to the sheet with the parameters specified.
        
        // How to add a word art to the worksheet specifying its properties, color, size, etc.
        
        // Insert a word art to the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let textProps = Api.CreateTextPr();
        textProps.SetFontSize(72);
        textProps.SetBold(true);
        textProps.SetCaps(true);
        textProps.SetColor(51, 51, 51, false);
        textProps.SetFontFamily("Comic Sans MS");
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(1 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
        worksheet.AddWordArt(textProps, "R7 Office", "textArchUp", fill, stroke, 0, 100 * 36000, 20 * 36000, 0, 2, 2 * 36000, 3 * 36000);
        
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
