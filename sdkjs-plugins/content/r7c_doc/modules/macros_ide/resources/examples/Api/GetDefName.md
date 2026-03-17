/**
 * @file GetDefName_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetDefName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve an `ApiName` object by its defined name.
 * It first adds a defined name "numbers" to the range A1:B1, then retrieves this defined name
 * using `Api.GetDefName()`, and displays its name in cell A3 of the active worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект `ApiName` по его определенному имени.
 * Сначала он добавляет определенное имя "numbers" к диапазону A1:B1, затем извлекает это определенное имя
 * с помощью `Api.GetDefName()` и отображает его имя в ячейке A3 активного листа.
 *
 * @param {string} name - The defined name to retrieve. (Определенное имя для получения.)
 * @returns {ApiName} The ApiName object representing the defined name. (Объект ApiName, представляющий определенное имя.)
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
        
        // Get the defined name object
        // Получение объекта определенного имени
        let defName = Api.GetDefName("numbers");
        
        // Display the name of the defined name in cell A3
        // Отображение имени определенного имени в ячейке A3
        worksheet.GetRange("A3").SetValue("DefName: " + defName.GetName());
        
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
