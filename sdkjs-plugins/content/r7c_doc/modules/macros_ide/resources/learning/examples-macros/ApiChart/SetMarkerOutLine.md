/**
 * @file SetMarkerOutLine_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiChart.SetMarkerOutLine
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the outline for a marker in a chart series.
 * It creates a scatter chart, sets its title, and then sets the outline for the first marker of both series.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить контур для маркера в ряду диаграммы.
 * Он создает точечную диаграмму, устанавливает ее заголовок, а затем устанавливает контур для первого маркера обоих рядов.
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
        // This example sets the outline to the marker in the specified chart series.
        
        // Create the "scatter" chart and set an outline of the specified width and color to its markers.
        
        // How to use the ApiStroke object as an outline of the chart markers.
        
        // How to outline the markers of the ApiChart object.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("B1").SetValue(2014);
        worksheet.GetRange("C1").SetValue(2015);
        worksheet.GetRange("D1").SetValue(2016);
        worksheet.GetRange("A2").SetValue("Projected Revenue");
        worksheet.GetRange("A3").SetValue("Estimated Costs");
        worksheet.GetRange("B2").SetValue(200);
        worksheet.GetRange("B3").SetValue(250);
        worksheet.GetRange("C2").SetValue(240);
        worksheet.GetRange("C3").SetValue(260);
        worksheet.GetRange("D2").SetValue(280);
        worksheet.GetRange("D3").SetValue(280);
        let chart = worksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "scatter", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);
        chart.SetTitle("Financial Overview", 13);
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
        chart.SetMarkerFill(fill, 0, 0, true);
        let stroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
        chart.SetMarkerOutLine(stroke, 0, 0, true);
        fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        chart.SetMarkerFill(fill, 1, 0, true);
        chart.SetMarkerOutLine(stroke, 1, 0, true);
        
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
