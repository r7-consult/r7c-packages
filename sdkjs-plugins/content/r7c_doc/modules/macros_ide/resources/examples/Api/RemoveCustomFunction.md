/**
 * @file RemoveCustomFunction_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.RemoveCustomFunction
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to remove a custom function from the R7 Office API.
 * It first adds a custom function named "ADD" to a library, then uses it in cell A1,
 * and finally removes the custom function, displaying a message in cell A3.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как удалить пользовательскую функцию из API R7 Office.
 * Сначала он добавляет пользовательскую функцию с именем "ADD" в библиотеку, затем использует ее в ячейке A1,
 * и, наконец, удаляет пользовательскую функцию, отображая сообщение в ячейке A3.
 *
 * @param {string} functionName - The name of the custom function to remove. (Имя пользовательской функции для удаления.)
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
        
        // Add a custom function library and a custom function named ADD
        // Добавление библиотеки пользовательских функций и пользовательской функции с именем ADD
        Api.AddCustomFunctionLibrary("LibraryName", function(){
            /**
             * Function that returns the argument
             * Функция, которая возвращает аргумент
             * @customfunction
             * @param {any} first First argument. (Первый аргумент.)
             * @returns {any} second Second argument. (Второй аргумент.)
             */
            Api.AddCustomFunction(function ADD(first, second) {
                return first + second;
            });
        });
        
        // Get the active worksheet
        // Получение активного листа
        let worksheet = Api.GetActiveSheet();
        
        // Use the custom function in cell A1
        // Использование пользовательской функции в ячейке A1
        worksheet.GetRange("A1").SetValue("=ADD(1, 2)");
        
        // Remove the custom function
        // Удаление пользовательской функции
        Api.RemoveCustomFunction("add");
        
        // Display a message indicating the custom function was removed
        // Отображение сообщения о том, что пользовательская функция была удалена
        worksheet.GetRange("A3").SetValue("The ADD custom function was removed.");
        
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
