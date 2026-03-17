/**
 * @file onWorksheetChange_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.onWorksheetChange
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to attach and handle the `onWorksheetChange` event in the R7 Office API.
 * It attaches an event listener that logs a message and the address of the changed range to the console
 * whenever a change occurs on the worksheet. It then triggers a change by setting a value in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как прикрепить и обработать событие `onWorksheetChange` в API R7 Office.
 * Он прикрепляет слушатель событий, который выводит сообщение и адрес измененного диапазона в консоль
 * при каждом изменении на листе. Затем он вызывает изменение, устанавливая значение в ячейке A1.
 *
 * @param {object} range - The range object that was changed. (Объект диапазона, который был изменен.)
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
        
        // Attach an event listener for onWorksheetChange
        // Прикрепление слушателя событий для onWorksheetChange
        Api.attachEvent("onWorksheetChange", function(range){
            console.log("onWorksheetChange");
            console.log(range.GetAddress());
        });
        
        // Get the active worksheet
        // Получение активного листа
        let worksheet = Api.GetActiveSheet();
        
        // Get range A1
        // Получение диапазона A1
        let range = worksheet.GetRange("A1");
        
        // Set a value in cell A1 to trigger the onWorksheetChange event
        // Установка значения в ячейке A1 для вызова события onWorksheetChange
        range.SetValue("1");
        
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
