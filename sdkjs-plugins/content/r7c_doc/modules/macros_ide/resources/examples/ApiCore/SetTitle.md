/**
 * @file SetTitle_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCore.SetTitle
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the title of the workbook.
 * It sets a title for the workbook and then displays it in a shape on the worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить заголовок рабочей книги.
 * Он устанавливает заголовок для рабочей книги, а затем отображает его в фигуре на листе.
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
        // This example demonstrates how to set the title of the workbook.
        
        const worksheet = Api.GetActiveSheet();
        const core = Api.GetCore();
        
        core.SetTitle("My Workbook Title");
        const title = core.GetTitle();
        
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(100, 50, 200));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        const shape = worksheet.AddShape(
        	"rect",
        	100 * 36000, 100 * 36000,
        	fill, stroke,
        	0, 0, 3, 0
        );
        
        let paragraph = shape.GetContent().GetElement(0);
        paragraph.AddText("Title: " + title);
        
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
