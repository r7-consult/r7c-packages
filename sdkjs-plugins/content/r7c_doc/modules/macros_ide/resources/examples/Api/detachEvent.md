/**
 * @file detachEvent_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.detachEvent
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to detach an event listener from the R7 Office API.
 * It first attaches an `onWorksheetChange` event, then immediately detaches it,
 * showing how to stop event handling.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как отсоединить слушатель событий от API R7 Office.
 * Сначала он прикрепляет событие `onWorksheetChange`, затем немедленно отсоединяет его,
 * показывая, как остановить обработку событий.
 *
 * @param {string} eventName - The name of the event to detach (e.g., "onWorksheetChange"). (Имя события для отсоединения.)
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
        
        // Set a value to trigger a potential change (though event will be detached)
        // Установка значения для вызова потенциального изменения (хотя событие будет отсоединено)
        range.SetValue("1");
        
        // Attach an event (for demonstration of detachment)
        // Прикрепление события (для демонстрации отсоединения)
        Api.attachEvent("onWorksheetChange", function(range){
            console.log("onWorksheetChange");
            console.log(range.GetAddress());
        });
        
        // Detach the event
        // Отсоединение события
        Api.detachEvent("onWorksheetChange");
        
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
