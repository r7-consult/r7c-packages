/**
 * @file attachEvent_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.attachEvent
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to attach an event listener to the R7 Office API.
 * It attaches an `onWorksheetChange` event listener that logs a message and the address of the changed range to the console
 * whenever a change occurs on the worksheet. It then triggers a change by setting a value in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как прикрепить слушатель событий к API R7 Office.
 * Он прикрепляет слушатель события `onWorksheetChange`, который выводит сообщение и адрес измененного диапазона в консоль
 * при каждом изменении на листе. Затем он вызывает изменение, устанавливая значение в ячейке A1.
 *
 * @param {string} eventName - The name of the event to attach (e.g., "onWorksheetChange"). (Имя события для прикрепления.)
 * @param {function} callback - The function to be executed when the event occurs. (Функция, которая будет выполнена при возникновении события.)
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
        
        // Get the active worksheet and a range
        // Получение активного листа и диапазона
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1");
        
        // Set a value in cell A1 to trigger the onWorksheetChange event
        // Установка значения в ячейке A1 для вызова события onWorksheetChange
        range.SetValue("1");
        
        // Attach an event listener for onWorksheetChange
        // Прикрепление слушателя событий для onWorksheetChange
        Api.attachEvent("onWorksheetChange", function(range){
            console.log("onWorksheetChange");
            console.log(range.GetAddress());
        });
        
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
