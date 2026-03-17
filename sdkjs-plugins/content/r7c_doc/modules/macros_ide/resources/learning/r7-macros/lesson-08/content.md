# Пример макроса: Выделение цветом ячеек с одинаковым значением


**Исходный пример на VBA**

```
Sub example()
    Dim xRg         As Range
    Dim xTxt        As String
    Dim xCell       As Range
    Dim xChar       As String
    Dim xCellPre    As Range
    Dim xCIndex     As Long
    Dim xCol        As Collection
    Dim I           As Long
    On Error Resume Next
    If ActiveWindow.RangeSelection.Count > 1 Then
        xTxt = ActiveWindow.RangeSelection.AddressLocal
    Else
        xTxt = ActiveSheet.UsedRange.AddressLocal
    End If
    Set xRg = Application.InputBox("please Select the data range:", "Kutools For Excel", xTxt, , , , , 8)
    If xRg Is Nothing Then Exit Sub
    xCIndex = 2
    Set xCol = New Collection
    For Each xCell In xRg
        On Error Resume Next
        xCol.Add xCell, xCell.Text
        If Err.Number = 457 Then
            xCIndex = xCIndex + 1
            Set xCellPre = xCol(xCell.Text)
            If xCellPre.Interior.ColorIndex = xlNone Then xCellPre.Interior.ColorIndex = xCIndex
            xCell.Interior.ColorIndex = xCellPre.Interior.ColorIndex
        ElseIf Err.Number = 9 Then
            MsgBox "Too many duplicate companies!", vbCritical, "Kutools For Excel"
            Exit Sub
        End If
        On Error GoTo 0
    Next
End Sub
```

**JavaScript Р7**

```
(function() {
    let whiteFill = Api.CreateColorFromRGB(255, 255, 255);
    let uniqueColorIndex = 0; // Текущий индекс в цветовом диапазоне

    let uniqueColors = [Api.CreateColorFromRGB(255, 255, 0),
        Api.CreateColorFromRGB(204, 204, 255),
        Api.CreateColorFromRGB(0, 255, 0),
        Api.CreateColorFromRGB(0, 128, 128),
        Api.CreateColorFromRGB(192, 192, 192),
        Api.CreateColorFromRGB(255, 204, 0)
    ]; // Массив с цветами

    function getColor() { // Функция, получающая цвета дубликатов 
        if (uniqueColorIndex === uniqueColors.length) {
            uniqueColorIndex = 0;
            a
        }
        return uniqueColors[uniqueColorIndex++];
    }

    let activeSheet = Api.ActiveSheet; // Получаем текущий лист
    let selection = activeSheet.Selection; // Получаем выделенную область
    let mapValues = {}; // Создаем пустой ассоциативный массив. В нем будет хранится информация о дубликатах.
    let arrRanges = []; //Массив всех клеток
    selection.ForEach(function(range) {

        let value = range.GetValue(); // Получаем значение из клеток
        if (!mapValues.hasOwnProperty(value)) {
            mapValues[value] = 0;
        }
        mapValues[value] += 1;
        arrRanges.push(range);
    });
    let value;
    let mapColors = {};
    //Окрашиваем дубликаты
    for (let i = 0; i < arrRanges.length; ++i) {
        value = arrRanges[i].GetValue();
        if (mapValues[value] > 1) {
            if (!mapColors.hasOwnProperty(value)) {
                mapColors[value] = getColor();
            }
            arrRanges[i].SetFillColor(mapColors[value]);
        } else {
            arrRanges[i].SetFillColor(whiteFill);
        }
    }
})();
```

Еще один пример того как , как можно выделить цветами одинаковые ячейки с данными

**JavaScript Р7**

```
/*
Макрос находит и выделяет одинаковым цветом все ячейки с одинаковым знаечнием. 
*/
(function(){ //Стандартная обёртка макроса
    let dictionary = []; //Массив-словарь всех уникальных значений ячеек в диапазоне поиска
    let area; //диапазон поиска
    
    //функция-обёртка для генерации трёх компонент цвета, рекурсивно генерирует цвет, пока не сгенериурет уникальный
    let rnd = function () {
        let color = Api.CreateColorFromRGB(Math.floor(Math.random() * 255), Math.floor(Math.random() * 255), Math.floor(Math.random() * 255)); //Генерируем случайный цвет
        if (!dictionary.some(item => item.value === color)){ //Если этого цвета нет в массиве-словаре
            return color; //То вернуть этот цвет
        } else {
            return rnd(); //Иначе рекурсивно попытаться сгенерировать другой и вернуть его
        }
    }
    if(Api.GetActiveSheet().GetSelection().GetCount()==1) { //Если выделена только одна ячейка
        area = Api.GetActiveSheet().GetUsedRange(); //То дипазон — весь лист
    } else {
        area = Api.GetActiveSheet().GetSelection(); //Иначе диапазон — выделенная область
    }
    
    area.ForEach((cell) => { //Перебираем каждую ячейку в диапазоне
        if (!dictionary.some(item => item.value === cell.GetValue())) //Если в массиве-словаре ещё нет значения равного текущей ячейке
            dictionary.push({   //То добавляем в конец массива-словаря объект с двумя полями
                value: cell.GetValue(), //Value — значение текущей ячейка
                color: rnd() //Color — случайно сгенерированный цвет маркирования
            });
    });
    area.ForEach((cell) => { //Перебираем каждую ячейку
        if(cell.GetValue().toString().length) //Если длина знаечния ячейки больше 0 (обрабатываем только не пустфе ячейки)
            cell.SetFillColor(dictionary[dictionary.findIndex(e => { //То установить цвет ячейки в соответствие полем Color 
                return e.value === cell.GetValue() //из элемента массива-словаря, совпадающего по значению Value
            })].color);
    });
})();
```




---
