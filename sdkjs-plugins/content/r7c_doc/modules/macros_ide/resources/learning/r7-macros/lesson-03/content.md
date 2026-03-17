# Практические примеры миграции VBA на JavaScript ч1


Преобразование макросов Excel VBA

Примеры:

Запись данных в ячейку листа

Цвет фона ячейки

Цвет шрифта ячейки

Жирный шрифт в ячейке

Объединение диапазона

Разъединение диапазона

Ширина столбца

Диапазон как таблица

Добавление диаграммы

Выделение дубликатов

Пустая строка

Вставка текста

Курсы валют

Импорт из CSV/TXT

Пересчет листа

Все строки и столбцы

Удаление фигур слайдов в презентации

Уникальный идентификатор

**Преобразование макросов Excel VBA**

**Макросы Р7-Офис** отличаются от макросов Microsoft, так как последние используют язык сценариев Visual Basic for Applications (VBA). JavaScript (язык макросов Р7) более гибок и может использоваться на любой платформе (что важно, так как редактор Р7 поддерживается на платформах Windows, Linux и Mac OS).  
  
Это может создать некоторые неудобства, если вы ранее использовали Microsoft Office с макросами VBA, так как они станут несовместимы с макросами Р7. Вы можете преобразовать ранее использованные макросы, чтобы использовать их с Р7-Офис.  
  
Процесс не слишком сложен. Рассмотрим следующий пример на языке VBA:

```
Sub Example()
    Dim myRange
    Dim result
    Dim Run         As Long
    
    For Run = 1 To 3
        Select Case Run
            Case 1
                result = "=SUM(A1:A100)"
            Case 2
                result = "=SUM(A1:A300)"
            Case 3
                result = "=SUM(A1:A25)"
        End Select
        ActiveSheet.range("B" & Run) = result
    Next Run
    
End Sub
```

Этот макрос считает сумму значений из трех диапазонов ячеек столбца A и помещает результаты в три ячейки столбца B.  
  
Того же самого можно достичь с помощью макросов Р7. Код будет практически идентичным и легко понятным, если вы знаете как VBA, так и JavaScript:

```
(function() {
    for (let run = 1; run <= 3; run++) {
        var result = "";
        switch (run) {
            case 1:
                result = "=SUM(A1:A100)";
                break;
            case 2:
                result = "=SUM(A1:A300)";
                break;
            case 3:
                result = "=SUM(A1:A25)";
                break;
            default:
                break;
        }
        Api.GetActiveSheet().GetRange("B" + run).Value = result;
    }
})();
```

**Почти любой код** макроса VBA (за некоторыми исключениями, описанными в нашем курсе) можно преобразовать в код на JavaScript, совместимый с макросами Р7.

**Запись данных в ячейку листа**

Записывает данные (фразу "Hello world") в ячейку третьего столбца четвертой строки рабочего листа.

**JavaScript Р7**

```
(function() {
    Api.GetActiveSheet().GetRange("C4").SetValue("Hello world");
})();
```

**Используемые методы**: GetActiveSheet, GetRange, SetValue

**VBA MS Office**

```
Sub example()
    Cells(4, 3) = "Hello world"
End Sub
```

**Изменение цвета фона ячейки**

Устанавливает синий цвет фона для ячейки B3.

**JavaScript Р7**

```
(function() {
    Api.GetActiveSheet().GetRange("B3").SetFillColor(Api.CreateColorFromRGB(0, 0, 250));
})();
```

Используемые методы: **GetActiveSheet**, **GetRange**, **SetFillColor**, **CreateColorFromRGB**

**VBA MS Office**

```
Sub example()
    Range("B3").Interior.Color = RGB(0, 0, 250)
End Sub
```

**Изменение цвета шрифта ячейки**

Устанавливает красный цвет шрифта для ячейки B4.

**JavaScript Р7**

```
(function() {
    Api.GetActiveSheet().GetRange("B4").SetFontColor(Api.CreateColorFromRGB(255, 0, 0));
})();
```

Используемые методы: **GetActiveSheet**, **GetRange**, **SetFontColor**

**VBA MS Office**

```
Sub example()
    Range("B4").Font.Color = RGB(255, 0, 0)
End Sub

```

**VBA MS Office**

```
Sub example()
    Range("B4").Font.Color = RGB(255, 0, 0)
End Sub
```

Установка жирного шрифта в ячейке

Устанавливает жирный шрифт для ячейки A2.

**JavaScript Р7**

```
(function() {
    Api.GetActiveSheet().GetRange("A2").SetBold(true);
})();
```

Используемые методы: **GetActiveSheet, GetRange, SetBold**

**VBA MS Office**

```
Sub example()
    Range("A2").Font.Bold = True
End Sub

```

Объединение диапазона ячеек

Объединяет выбранный диапазон ячеек.

**JavaScript Р7**

```
(function()
{
 Api.GetActiveSheet().GetRange("A1:B3").Merge(true);
})();

```

Используемые методы: **GetActiveSheet, GetRange, Merge**

**VBA MS Office**

```
Sub example()
Range("A1:B3").Merge
End Sub

```

Разъединение диапазона ячеек

Разъединяет выбранный диапазон ячеек.

**JavaScript Р7**

```
(function()
{
Api.GetActiveSheet().GetRange("C5:D10").UnMerge();
})();

```

Используемые методы: **GetActiveSheet, GetRange, UnMerge**

**VBA MS Office**

```
Sub example()
Range("C5:D10").UnMerge
End Sub

```

Установка ширины столбца B

Устанавливает ширину столбца

**JavaScript Р7**

```
(function()
{
 Api.GetActiveSheet().SetColumnWidth(1, 25);
})();

```

Используемые методы: **GetActiveSheet, SetColumnWidth**

**VBA MS Office**

```
Sub example()
Columns("B").ColumnWidth = 25
End Sub

```

Форматирование диапазона как таблицу

Форматирует диапазон ячеек A1:D10 как таблицу.

**JavaScript Р7**

```
(function()
{
Api.GetActiveSheet().FormatAsTable("A1:D10");
})();

```

Используемые методы: **GetActiveSheet, FormatAsTable**

**VBA MS Office**

```
Sub example()
Sheet1.ListObjects.Add(xlSrcRange, Range("A1:D10"), , xlYes).Name = "myTable1"
End Sub
```

Добавление диаграммы

Добавляет новую диаграмму в выбранный диапазон ячеек.

**JavaScript Р7**

```
((function()
{
Api.GetActiveSheet().AddChart("'Sheet1'!$C$5:$D$7", true, "bar", 2, 105 * 36000, 105 * 36000, 0, 0, 9, 0);
})();
```

Используемые методы: **GetActiveSheet, AddChart**

**VBA MS Office**

```
Sub example()
With ActiveSheet.ChartObjects.Add(Left:=300, Width:=300, Top:=10, Height:=300)
.Chart.SetSourceData Source:=Sheets("Sheet1").Range("C5:D7")
End With
End Sub
```

Нахождение следующей доступной пустой строки в рабочем листе

Этот макрос позволяет получить пустую строку в самом конце ваших данных (не между ними).

**JavaScript Р7**

```
(function ()
{
// Получаем активный лист
var activeSheet = Api.GetActiveSheet();
// Минимальный индекс строки
var indexRowMin = 0;
// Максимальный индекс строки
var indexRowMax = 1048576;
// Столбец 'A'
var indexCol = 0;
// Индекс строки для поиска пустой ячейки
var indexRow = indexRowMax;
for (; indexRow >= indexRowMin; --indexRow) {
// Получаем ячейку
var range = activeSheet.GetRangeByNumber(indexRow, indexCol);
// Проверяем значение
if (range.GetValue() && indexRow !== indexRowMax) {
range = activeSheet.GetRangeByNumber(indexRow + 1, indexCol);
range.Select();
break;
}
}
})();
```

Используемые методы: **GetActiveSheet, GetRangeByNumber, Select**

**VBA MS Office**

```
Sub example()
Range("A" & Rows.Count).End(xlUp).Offset(1).Select
End Sub
```

Вставка текста

Вставляет текст в документ в текущей позиции курсора.

**JavaScript Р7**

```
(function()
{
var oDocument = Api.GetDocument();
var oParagraph = Api.CreateParagraph();
oParagraph.AddText("Hello world!");
oDocument.InsertContent([oParagraph]);
})();
```

Используемые методы: **GetDocument, CreateParagraph, AddText, InsertContent**

**VBA MS Office**

```
Sub example()
Selection.TypeText Text:="Hello world!"
End Sub
```

**Курсы обмена валют**

Возвращает информацию о курсах обмена за последние несколько дней и заполняет таблицу полученными значениями. В этом макросе представлен пример для валютной пары USD-EUR, но вы можете получить информацию о других курсах обмена, изменив значение переменной `sCurPair` ("EUR\_USD", "BTC\_USD" и т.д.).  
  
**В этом макросе** используется сторонний сервис CurrencyConverterApi.com для получения информации о курсах обмена. Существует ограничение на количество запросов в час. Если этот лимит превышен, макрос не будет работать. Если вы хотите использовать этот макрос, лучше зарегистрироваться на сайте сервиса и использовать свой собственный ключ в коде макроса.  
  
Вы можете назначить этот макрос автофигуре. При нажатии на неё макрос выполнится, таблица заполнится актуальными данными и соответствующий график обновится.

**JavaScript Р7**

```
(function()
{
var sCurPair = "USD_EUR";

function formatDate(d) {
var month = '' + (d.getMonth() + 1),
day = '' + d.getDate(),
year = d.getFullYear();

if (month.length < 2)
month = '0' + month;
if (day.length < 2)
day = '0' + day;

return [year, month, day].join('-');
}

function previousWeek(){
var today = new Date();
var prevweek = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 7);
return prevweek;
}

var sDate = formatDate(previousWeek());
var sEndDate = formatDate(new Date());
var apiKey = 'ваш_ключ_от_API';
var sUrl = 'https://free.currconv.com/api/v7/convert?q='
+ sCurPair + '&compact=ultra' + '&date=' + sDate + "&endDate=" + sEndDate + '&apiKey=' + apiKey;
var xmlHttp = new XMLHttpRequest();
xmlHttp.open("GET", sUrl, false);
xmlHttp.send();
if (xmlHttp.readyState == 4 && xmlHttp.status == 200) {
var oData = JSON.parse(xmlHttp.responseText);
for(var key in oData) {
var sheet = Api.GetSheet("Sheet1");
var oRange = sheet.GetRangeByNumber(0, 1);
oRange.SetValue(key);
var oDates = oData[key];
var nRow = 1;
for(var date in oDates) {
oRange = sheet.GetRangeByNumber(nRow, 0);
oRange.SetValue(date);
oRange = sheet.GetRangeByNumber(nRow, 1);
oRange.SetValue(oDates[date]);
nRow++;
}
}
}
})();
```

Используемые методы: **GetSheet, GetRangeByNumber, SetValue**

**VBA MS Office**

```
Sub example()
' ' Этот пример не имеет прямого аналога в VBA и требует стороннего API
End Sub
```

Пересчет значений рабочего листа

Повторно пересчитывает значения ячеек рабочего листа с интервалом в одну секунду.

**JavaScript Р7**

```
(function ()
{
let timerId = setInterval(function(){
Api.asc_calculate(Asc.c_oAscCalculateType.All);
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("F2").SetValue("Average");
oWorksheet.GetRange("G2").SetValue("=AVERAGE(B2:B80)");
}, 1000);
})();
```

Используемые методы: **GetActiveSheet, GetRange, SetValue**

**VBA MS Office**

```
Sub example()
  Application.OnTime Now + TimeValue("00:00:01"), "RecalculateValues"
End Sub

Sub RecalculateValues()
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Sheets("Sheet1")
  
  ws.Calculate
  ws.Range("F2").Value = "Average"
  ws.Range("G2").Formula = "=AVERAGE(B2:B80)"
  
  ' 'Reschedule the recalculation
  Application.OnTime Now + TimeValue("00:00:01"), "RecalculateValues"
End Sub

```

**Отображение всех строк и столбцов**

Отображает все строки и столбцы в активной электронной таблице.

**JavaScript Р7**

```
(function()
{
var activeSheet = Api.GetActiveSheet();
var indexRowMax = 1048576;
var indexColMax = 16384; // Максимальное количество столбцов в Excel

// Отобразить все строки
for (let i = 0; i < indexRowMax; i++) {
activeSheet.GetRows(i + 1).SetHidden(false);
}

// Отобразить все столбцы
for (let j = 0; j < indexColMax; j++) {
activeSheet.GetColumns(j + 1).SetHidden(false);
}

var newRange = activeSheet.GetRange("A1");
newRange.SetValue("All the rows and columns are unhidden now");
})();
```

Используемые методы: **GetActiveSheet, GetRows, SetHidden, GetColumns, GetRange, SetValue**

**VBA MS Office**

```
Sub UnhideAllRowsAndColumns()
Cells.EntireRow.Hidden = False
Cells.EntireColumn.Hidden = False
Range("A1").Value = "All the rows and columns are unhidden now"
End Sub
```

Удаление фигур слайдов в презентации.

Удаляет фигуры слайдов в презентации.

**JavaScript Р7**

```
(function()
{
var oPresentation = Api.GetPresentation();
var slideCount = oPresentation.GetSlideCount(); // Получаем количество слайдов в презентации

for (let i = 0; i < slideCount; i++) {
var oSlide = oPresentation.GetSlideByIndex(i);
var aShapes = oSlide.GetAllShapes();

for (let j = 0; j < aShapes.length; j++) {
aShapes[j].Delete();
}
}
})();
```

Используемые методы: **GetPresentation, GetSlideByIndex, GetAllShapes, GetSlideCount, Delete**

**VBA MS Office**

```
Sub RemoveShapes()
Dim slide As slide
Dim shape As shape
For Each slide In ActivePresentation.Slides
For Each shape In slide.Shapes
shape.Delete
Next shape
Next slide
End Sub
```

Вставка уникального идентификатора

Вставляет уникальный идентификатор.

**JavaScript Р7**

```
(function()
{
function generate() {
let key = '';
const data = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ' +
'abcdefghijklmnopqrstuvwxyz0123456789';
for (let i = 1; i <= 12; i++) {
let index = Math.floor(Math.random() * data.length);
key += data.charAt(index);
}
return key;
}

const id = generate();
const oDocument = Api.GetDocument();
const oParagraph = Api.CreateParagraph();
oParagraph.AddText(id);
oDocument.InsertContent([oParagraph], { "KeepTextOnly": true });
})();
```

Используемые методы: **GetDocument, CreateParagraph, AddText, InsertContent**

**VBA MS Office**

```
Sub InsertUniqueId()
Dim uniqueId As String
uniqueId = GenerateUniqueId()

Sub InsertUniqueId()
  Dim uniqueId As String
  uniqueId = GenerateUniqueId()
  
  Dim para As Paragraph
  Set para = ActiveDocument.Paragraphs.Add
  para.Range.Text = uniqueId
End Sub

Function GenerateUniqueId() As String
  Dim i As Integer
  Dim key As String
  Dim data As String
  
  data = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
  
  For i = 1 To 12
    key = key & Mid(data, Int((Len(data) * Rnd) + 1), 1)
  Next i
  
  GenerateUniqueId = key
End Function
```




---
