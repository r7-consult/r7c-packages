/**
 * @file ChangeChartType_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiChartSeries.ChangeChartType
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to change the chart type of a series.
 * It creates a combo chart, gets the first series, changes its type to "area", and displays the old and new types in the worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как изменить тип диаграммы для ряда.
 * Он создает комбинированную диаграмму, получает первый ряд, изменяет его тип на "area" и отображает старый и новый типы на листе.
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
        // This example changes the type of the first series of ApiChart class and inserts the new type into the document.
        
        // How to change a chart type to an area one.
        
        // Change a chart type.
        
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
        let chart = worksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "comboBarLine", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);
        chart.SetTitle("Financial Overview", 13);
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
        chart.SetSeriesFill(fill, 0, false);
        fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        chart.SetSeriesFill(fill, 1, false);
        let series = chart.GetSeries(0);
        let seriesType = series.GetChartType();
        worksheet.GetRange("F1").SetValue("Old Series Type = " + seriesType);
        series.ChangeChartType("area");
        seriesType = series.GetChartType();
        worksheet.GetRange("F2").SetValue("New Series Type = " + seriesType);
        
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
