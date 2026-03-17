/**
 * @file RecalculateAllFormulas_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.RecalculateAllFormulas
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to recalculate all formulas in the active workbook.
 * It sets values in cells B1 and C1, then sets formulas in A1 and E1 that depend on these values.
 * After changing the value in B1, it calls `Api.RecalculateAllFormulas()` to update the formula results,
 * and displays a message in A3.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как пересчитать все формулы в активной книге.
 * Он устанавливает значения в ячейках B1 и C1, затем устанавливает формулы в A1 и E1, которые зависят от этих значений.
 * После изменения значения в B1 он вызывает `Api.RecalculateAllFormulas()` для обновления результатов формул,
 * и отображает сообщение в A3.
 *
 * @returns {void}
 *
 * @see https://r7-consult.ru/
 */

(function() {
    'use strict';
    
    try {
        // Initialize R7 Office API
        // Инициализация API R7 Office
        const api = Api;
        if (!api) {
            throw new Error('R7 Office API not available');
        }
        
        // Get the active worksheet
        // Получение активного листа
        let worksheet = Api.GetActiveSheet();
        
        // Set initial values in cells B1 and C1
        // Установка начальных значений в ячейках B1 и C1
        worksheet.GetRange("B1").SetValue(1);
        worksheet.GetRange("C1").SetValue(2);
        
        // Set a formula in cell A1 that sums B1 and C1
        // Установка формулы в ячейке A1, которая суммирует B1 и C1
        let rangeA1 = worksheet.GetRange("A1");
        rangeA1.SetValue("=SUM(B1:C1)");
        
        // Set another formula in cell E1 that adds 1 to A1
        // Установка другой формулы в ячейке E1, которая добавляет 1 к A1
        let rangeE1 = worksheet.GetRange("E1");
        rangeE1.SetValue("=A1+1");
        
        // Change the value in B1 to trigger recalculation
        // Изменение значения в B1 для запуска пересчета
        worksheet.GetRange("B1").SetValue(3);
        
        // Recalculate all formulas in the workbook
        // Пересчет всех формул в книге
        Api.RecalculateAllFormulas();
        
        // Display a message indicating that formulas were recalculated
        // Отображение сообщения о том, что формулы были пересчитаны
        worksheet.GetRange("A3").SetValue("Formulas from cells A1 and E1 were recalculated with a new value from cell C1.");
        
        // Success notification
        // Уведомление об успешном выполнении
        console.log('Macro executed successfully');
        
    } catch (error) {
        // Error handling
        // Обработка ошибок
        console.error('Macro execution failed:', error.message);
        // Optional: Show error to user in cell A1 if API is available
        // Опционально: Показать ошибку пользователю в ячейке A1, если API доступен
        if (typeof Api !== 'undefined' && Api.GetActiveSheet) {
            const sheet = Api.GetActiveSheet();
            if (sheet) {
                sheet.GetRange('A1').SetValue('Error: ' + error.message);
            }
        }
    }
})();
