# Пример макроса: Миграция скрипта VBA по объединению данных из листов в одну таблицу в R7 JavaScript


Пусть у нас есть книга с большим количеством листов с данными, разбитыми по одним и тем же столбцам, например набор отчетов о заказах от менеджеров в фирме за месяц, и требуется объединить данные в одну таблицу, чтобы анализировать данные по всем заказам.  
  
Вручную это делать долго и сложно, поэтому решим задачу сначала с помощью макроса VBA для Excel, а потом преобразуем этот макрос в Javascript для Р7-Офис.  
  
При написании макроса проверим наличие в книге листов с исходными данными. Если их нет, прекратим работу макроса. Если листы с данными имеются, создадим лист AllData для объединенных данных, если его нет. При его наличии очистим уже имеющиеся данные на нем. На лист AllData нужно один раз скопировать названия столбцов с любого листа с данными. После этого скопируем данные из листов с исходными данными.  
  
В дальнейшем может понадобиться копирование данных в отдельный файл. Можно было бы сделать вариант макроса с параметрами файла для вывода результата в виде опций, но пока оставим дополнительный вариант на будущее.

Макрос на VBA

```
Sub CollectDataFromAllSheets()
   Dim ws As Worksheet
   Dim allDataSheet As Worksheet
   Dim rngData As Range
   Dim currentResultRow As Long
   Dim sheetExists As Boolean
   Dim sheetName As String

   sheetName = "AllData"
   sheetExists = False

   ' Проверка наличия листа с именем "AllData"
   For Each ws In ActiveWorkbook.Worksheets
       If ws.Name = sheetName Then
sheetExists = True
           Set allDataSheet = ws
           Exit For
       End If
   Next ws

   ' Если лист "AllData" найден, очищаем его, иначе создаем новый
   If sheetExists Then
allDataSheet.Cells.Clear
   Else
       Set allDataSheet = ActiveWorkbook.Worksheets.Add
allDataSheet.Name = sheetName
   End If

   ' Копируем заголовок с первого листа
Worksheets(1).Rows(1).Copy Destination:=allDataSheet.Rows(1)

currentResultRow = 2

   ' Перебор всех листов для копирования данных
   For Each ws In ActiveWorkbook.Worksheets
       If ws.Name <> sheetName Then
           Set rngData = ws.Range("A2", ws.Range("A2").SpecialCells(xlCellTypeLastCell)) ' от A2 до последней ячейки

rngData.Copy Destination:=allDataSheet.Cells(currentResultRow, 1)
currentResultRow = currentResultRow + rngData.Rows.Count
       End If
   Next ws
End Sub
```

Макрос на JavaScript для Р7-Офис

```
(function()
{
    //Вспомогательная функция, находит последний используемый ряд в данном листе. В отличии от остальных методов, не останавливается на пустых промежутках.
function FindLastRow(sheet) 
{
    
    // Minimum row index
    let indexRowMin = 0;
    // Maximum row index
    let indexRowMax = 1048576;
    // Column 'A'
    let indexCol = 0;
    // Row index for empty cell search
    let indexRow = indexRowMax;
    for (; indexRow >= indexRowMin; --indexRow) {
        // Getting the cell
        var range = sheet.GetRangeByNumber(indexRow, indexCol);
        // Checking the value
        if (range.GetValue() && indexRow !== indexRowMax) {
            //Нашли первое значение, считая снизу, возвращаем его
			return indexRow;
        }
    }
}

function CollectDataFromAllSheets() {
	
	
	//Запоминаем листы, из которых нам надо собрать данные и добавляем лист для объединенных данных
	let worksheets = Api.GetSheets();
	if (worksheets.count == 0) {
		return;
	}

    Api.AddSheet();
    let allDataSheet = Api.ActiveSheet;
	//Если требуется сохранять не в отдельный лист, а в отдельную книгу - то вместо предыдущей строчки используем:
	/*
	builder.CreateFile("xlsx");
	let allDataSheet = Api.AddSheet();
	*/

	 //копируем в итоговый лист заголовки столбцов из первого листа
	 worksheets[1].GetRows("1").Copy(allDataSheet.GetRows("1"));
    
	
	//И устанавливаем текущую строку равной 2(сразу после заголовков)
	let currentResultRow = 2;

	//проходим по всем листам, из которых надо собрать данные
	for (let ws of worksheets) {
		//Для каждого листа находим последний ряд
		let lastRow = FindLastRow(ws) +1;
		//console.log("LastRow: ", lastRow);
		//И копируем каждый ряд в конечную таблицу
		 for (let rowNum = 2; rowNum <= lastRow; rowNum++) {
		     //-2 потому что rowNum у нас тоже идет от 2
		    ws.GetRows(rowNum).Copy(allDataSheet.GetRows(rowNum + currentResultRow-2));
		    
		 }
	
        //После прохода по всем рядам - обновляем currentResultRow (можно и с помощью FindLastRow, но это будет значительно более затратно по вычислениям)
        let rowCount = lastRow-1;
		currentResultRow += rowCount;
         

	}
	
}

CollectDataFromAllSheets();

})();
```




---
