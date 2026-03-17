/**
 * @file AddCustomFunction_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.AddCustomFunction
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to add a custom function to the R7 Office API and use it in a spreadsheet.
 * It defines a custom function named "ADD" that sums two arguments, adds it to a custom function library,
 * and then uses this custom function in cell A1 of the active worksheet to calculate the sum of 1 and 2.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как добавить пользовательскую функцию в API R7 Office и использовать ее в электронной таблице.
 * Он определяет пользовательскую функцию с именем "ADD", которая суммирует два аргумента, добавляет ее в библиотеку пользовательских функций,
 * а затем использует эту пользовательскую функцию в ячейке A1 активного листа для вычисления суммы 1 и 2.
 *
 * @param {function} customFunction - The custom function to add. (Пользовательская функция для добавления.)
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
        worksheet.GetRange('A1').SetValue('=ADD(1,2)');
        
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
