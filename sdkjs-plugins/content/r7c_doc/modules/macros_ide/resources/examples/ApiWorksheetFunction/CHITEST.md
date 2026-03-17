/**
 * @file CHITEST_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.CHITEST
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the test for independence: the value from the chi-squared distribution for the statistic and the appropriate degrees of freedom.
 * It sets up actual and expected data, and then calculates the chi-squared test statistic and displays the result in cell D6.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть тест на независимость: значение из распределения хи-квадрат для статистики и соответствующие степени свободы.
 * Он настраивает фактические и ожидаемые данные, а затем вычисляет статистику теста хи-квадрат и отображает результат в ячейке D6.
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
        // This example shows how to return the test for independence: the value from the chi-squared distribution for the statistic and the appropriate degrees of freedom.
        
        // How to return the value from the chi-squared distribution for the statistic and the appropriate degrees of freedom.
        
        // Use function to return the value from the chi-squared distribution for the statistic and the appropriate degrees of freedom.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let actual1 = ["Actual", 58, 11, 10];
        let actual2 = ["Actual", 35, 25, 23];
        let expected1 = ["Expected", 45.35, 17.56, 16.09];
        let expected2 = ["Expected", 47.65, 18.44, 16.91];
        
        for (let i = 0; i < actual1.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(actual1[i]);
        }
        for (let j = 0; j < actual2.length; j++) {
            worksheet.GetRange("B" + (j + 1)).SetValue(actual2[j]);
        }
        for (let n = 0; n < expected1.length; n++) {
            worksheet.GetRange("C" + (n + 1)).SetValue(expected1[n]);
        }
        for (let m = 0; m < expected2.length; m++) {
            worksheet.GetRange("D" + (m + 1)).SetValue(expected2[m]);
        }
        
        let range1 = worksheet.GetRange("A2:B4");
        let range2 = worksheet.GetRange("C2:D4");
        worksheet.GetRange("D6").SetValue(func.CHITEST(range1, range2));
        
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
