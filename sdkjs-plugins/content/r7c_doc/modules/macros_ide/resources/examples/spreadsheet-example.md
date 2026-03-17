/*********************
 * P7 Офис Spreadsheet Macro Example
 * 
 * Этот макрос демонстрирует основные операции с электронной таблицей:
 * - Получение активного листа
 * - Выбор диапазона ячеек
 * - Настройка цвета фона
 * - Отображение сообщения
 
 -------------
 Lad — центр компетенций по импортозамещению
 
   Помогаем перейти на импортонезависимые программные продукты, 
входящие в реестр российского ПО. 
  Внедряем и сопровождаем отечественное программное обеспечение, 
обучаем пользователей и администраторов 

https://lad-soft.ru/
gov@lad24.ru
--------------

***********************/

(function() {
    // Get the active spreadsheet
    var oSheet = Api.GetActiveSheet();
    
    // Select range A1:J20
    var oRange = oSheet.GetRange("A1:J20");
    
    // Set background color to red
    oRange.SetFillColor(255, 0, 0);
    
    // Show completion message  
    Api.ShowMessage("Spreadsheet Macro", "Cells A1:J20 have been filled with red color.");
    
    // This line will cause an error for demonstration
    // UnknownApi.DoSomething();
    
    console.log("Spreadsheet macro executed successfully!");
})();
