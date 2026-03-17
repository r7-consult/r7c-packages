/**
 * @file AddDefName_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.AddDefName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to add a new defined name to a range of cells in the spreadsheet.
 * It sets values in cells A1 and B1, then defines a name "numbers" for the range "Sheet1!$A$1:$B$1",
 * and displays a confirmation message in cell A3.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как добавить новое определенное имя к диапазону ячеек в электронной таблице.
 * Он устанавливает значения в ячейках A1 и B1, затем определяет имя "numbers" для диапазона "Sheet1!$A$1:$B$1",
 * и отображает сообщение о подтверждении в ячейке A3.
 *
 * @param {string} name - The name to define for the range. (Имя для определения диапазона.)
 * @param {string} reference - The range reference (e.g., "Sheet1!$A$1:$B$1"). (Ссылка на диапазон.)
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
        
        // Set values in cells A1 and B1
        // Установка значений в ячейках A1 и B1
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        
        // Add a defined name "numbers" to the range A1:B1
        // Добавление определенного имени "numbers" к диапазону A1:B1
        Api.AddDefName("numbers", "Sheet1!$A$1:$B$1");
        
        // Display a message confirming the defined name creation
        // Отображение сообщения, подтверждающего создание определенного имени
        worksheet.GetRange("A3").SetValue("We defined a name 'numbers' for a range of cells A1:B1.");
        
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
