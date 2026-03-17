# Практические примеры миграции VBA на JavaScript ч3


Завершение адресной информации

Граница ячейки

Последняя ячейка в столбце

Разные диаграммы

Выделение дубликатов цветом в области

Числовой тип

Ячейки из диапазона

Текст из ячейки

Значение из области в ячейку

Создание параграфа

Цвета фона диапазона

Жирность для диапазона

Цвет шрифта диапазона

Слияние диапазона ячеек

Диапазон для сортировки

Разъединение диапазона

Номера строк из диапазона

Ширина столбца

Очистка диапазона

Шрифт страницы

Список ключей localStorage

Удаление ключа из localStorage

Установка ключа в LocalStorage

Завершение адресной информации

Дополняет основные данные адреса подробной адресной информацией и вставляет её в электронную таблицу.

**JavaScript Р7**

```
// Структура макроса
// 1. Считывание адреса (ячейка A2)
// 2. Отправка запроса
// 3. Получение ответа и создание объекта с подробной адресной информацией
// 4. Вставка подробной информации об адресе
// 5. Считывание адреса в следующей строке (ячейка A3) и отправка запроса

(function()
{
const API_KEY = 'yourAPIkey';
const ENDPOINT = 'https://api.geoapify.com/v1/geocode/search';
const oWorksheet = Api.GetActiveSheet();
let row = 2;
makeRequest(oWorksheet.GetRange(`A${row}`).GetText());

// Отправка запроса
function makeRequest(ADDRESS) {
if (ADDRESS === '') return;
$.ajax({
url: `${ENDPOINT}?text=${addressToRequest(ADDRESS)}&apiKey=${API_KEY}`,
dataType: 'json',
}).done(successFunction);
}

// Преобразование адреса для запроса (London, United Kingdom -> London%2C%20United%20Kingdom)
function addressToRequest(address) {
return address.replaceAll(' ', '%20').replaceAll(',', '%2C');
}

// Обработка ответа
function successFunction(response) {
const data = createAddressDetailsObject(response);
pasteAddressDetails(data);
reload();
}

// Создание объекта с подробной информацией об адресе, если адрес найден
function createAddressDetailsObject(response) {
if (response.features.length === 0) {
return { error: 'Address not found' };
}
console.log(response);
let data = {
country: response.features[0].properties.country,
county: response.features[0].properties.county,
city: response.features[0].properties.city,
post_code: response.features[0].properties.postcode,
full_address_line: response.features[0].properties.formatted
};
data = checkMissingData(data);
return data;
}

// Замена отсутствующих полей на '-'
function checkMissingData(data) {
Object.keys(data).forEach(key => {
if (data[key] === undefined) data[key] = '-';
});
return data;
}

// Вставка подробной информации об адресе
function pasteAddressDetails(data) {
const oRange = oWorksheet.GetRange(`B${row}:F${row}`);
if (data.error !== undefined) {
oRange.SetValue([[data.error]]);
} else {
oRange.SetValue([
[
data.country,
data.county,
data.city,
data.post_code,
data.full_address_line
]
]);
}
// Выполнение рекурсивно до тех пор, пока значение "Address" не станет пустым
row++;
makeRequest(oWorksheet.GetRange(`A${row}`).GetText());
}

// Обновление листа при изменениях
function reload() {
let reload = setInterval(function(){
Api.asc_calculate(Asc.c_oAscCalculateType.All);
});
}
})();
```

Используемые методы: GetActiveSheet, GetRange, SetValue, GetText

**VBA MS Office**

```
Sub example()
'' Этот пример не имеет прямого аналога на VBA
End Sub
```

Граница ячейки

Функция, получающая ячейку и отрисовывающая границу

**JavaScript Р7**

```
function aroundWhiteBorder(el) {
 el.SetBorders("Bottom", "Medium", Api.CreateColorFromRGB(255, 255, 255));
 el.SetBorders("Top", "Medium", Api.CreateColorFromRGB(255, 255, 255));
 el.SetBorders("Left", "Medium", Api.CreateColorFromRGB(255, 255, 255));
 el.SetBorders("Right", "Medium", Api.CreateColorFromRGB(255, 255, 255));
}
```

Функция, принимающая индекс столбца и выделяющая последнюю ячейку   
в нем

Выделение последней ячейки в столбце по его индексу

**JavaScript Р7**

```
//Перемещение в конец выбранного столбца
(function () 
{
 let activeSheet = Api.ActiveSheet; // Получение текущего листа
 let indexRowMin = 0; // Минимальный индекс строки
 let indexRowMax = 1048576; // Максимальный индекс строки

 let indexCol = 0; // Индекс нужного столбца

 let indexRow = indexRowMax;
 for (; indexRow >= indexRowMin; --indexRow) {
 let range = activeSheet.GetRangeByNumber(indexRow, indexCol);
 if (range.GetValue() && indexRow !== indexRowMax) {
 range = activeSheet.GetRangeByNumber(indexRow + 1, indexCol);
 range.Select();
 break;
 }
 }
})();
```

Помещение в документ диаграммы разного типа

Помещение в документ диаграммы разного типа

**JavaScript Р7**

```
//Создание диаграмм в текстовом редакторе
(function()
{
let oDocument = Api.GetDocument();
let oParagraph = oDocument.GetElement(0);

let sType = "bar3D" //Тип диаграммы / "bar" | "barStacked" | "barStackedPercent" | "bar3D" | "barStacked3D" | "barStackedPercent3D" | "barStackedPercent3DPerspective" | "horizontalBar" | "horizontalBarStacked" | "horizontalBarStackedPercent" | "horizontalBar3D" | "horizontalBarStacked3D" | "horizontalBarStackedPercent3D" | "lineNormal" | "lineStacked" | "lineStackedPercent" | "line3D" | "pie" | "pie3D" | "doughnut" | "scatter" | "stock" | "area" | "areaStacked" | "areaStackedPercent"

let aSeries = [[200, 240, 280],[250, 260, 280]] //Массив данных
let aSeriesNames = ["Projected Revenue", "Estimated Costs"]//Массив имён данных
let aCatNames = [2014, 2015, 2016] //Массив имён категорий
let width = 4051300 // Ширина
let height = 2347595 // Длина
let styleIndex = 24 // Индекс стиля диаграммы по спецификации OOXML(1 - 48)
let aNumFormats = ["0", "0.00"]

let oDrawing = Api.CreateChart(sType, aSeries, aSeriesNames, aCatNames, width, height, styleIndex, aNumFormats);
oDrawing.SetShowPointDataLabel(1, 1, false, false, true, false); // Создание объекта для отрисовки / Индекс значения из массива, над которым будет значение - int/ Индекс столбца, над которым будет значение - int/ Демонстрация имён таблицы - bool/ Демонстрация строк таблицы - bool/ Демонстрация значения данных диаграммы - bool/ Демонстрация процента значений данных - bool
oParagraph.AddDrawing(oDrawing);
})();
```

Выделение дубликатов разными цветами в выбранной области

Выделение дубликатов разными цветами в выбранной области

**JavaScript Р7**

```
(function () 
{
 let whiteFill = Api.CreateColorFromRGB(255, 255, 255);
 let uniqueColorIndex = 0; // Текущий индекс в цветовом диапазоне
 
 let uniqueColors = [Api.CreateColorFromRGB(255, 255, 0),
 Api.CreateColorFromRGB(204, 204, 255),
 Api.CreateColorFromRGB(0, 255, 0),
 Api.CreateColorFromRGB(0, 128, 128),
 Api.CreateColorFromRGB(192, 192, 192),
 Api.CreateColorFromRGB(255, 204, 0)]; // Массив с цветами

 function getColor() { // Функция, получающая цвета дубликатов 
 if (uniqueColorIndex === uniqueColors.length) {
 uniqueColorIndex = 0;a
 }
 return uniqueColors[uniqueColorIndex++];
 }

 let activeSheet = Api.ActiveSheet; // Получаем текущий лист
 let selection = activeSheet.Selection; // Получаем выделенную область
 let mapValues = {}; // Создаем пустой ассоциативный массив. В нем будет хранится информация о дубликатах.
 let arrRanges = []; //Массив всех клеток
 selection.ForEach(function (range) {
 
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

Приведение ячейки к числовому типу

Приведение ячейки к числовому типу данных

**JavaScript Р7**

```
(function()
{
let oWorksheet = Api.GetActiveSheet();// Получаем текущий лист
let test = oWorksheet.Selection; // Объявляем переменную = ее к отбору по ячейке
test.ForEach(x => { // Создаем Цикл
let value = x.GetValue(); // Объявляем переменную = ее к замене
if(value === null || value === "" || !Number(value)){
return;
}
else{
value = Number(value); // переменная приравнивается к числовому значению
x.SetValue(value); // выводим переменную
x.SetNumberFormat("0.00"); // изменяем формат ячейки на числовой и форматируем под определенный вид числа
}
});
})();
```

Получение ячеек из диапазона

Получение ячеек из диапазона

**JavaScript Р7**

```
let oWorksheet = Api.GetActiveSheet(); //Получение текущего листа
let oRange = oWorksheet.GetRange("A1:C3"); //Получение диапазона ячеек
oRange.GetCells(2, 1).SetFillColor(Api.CreateColorFromRGB(255, 224, 204)); //Получение ячеек из диапазона
```

Получение текста из ячейки

Получение текста из ячейки

**JavaScript Р7**

```
let oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("text1");
oWorksheet.GetRange("B1").SetValue("text2");
oWorksheet.GetRange("C1").SetValue("text3");
let oRange = oWorksheet.GetRange("A1:C1");
let sText = oRange.GetText();
oWorksheet.GetRange("A3").SetValue("Text from the cell A1: " + sText);
```

Значение из выбранной области в ячейку

Получение значения из выбранной области и вставка в нужную ячейку

**JavaScript Р7**

```
(function()
{
 let oWorksheet = Api.GetActiveSheet();// Получаем текущий лист
 let test = oWorksheet.Selection; // Объявляем переменную = ее к отбору по ячейке
 test.ForEach(x => {
 let value = x.GetValue();
 
 oWorksheet.GetRange("M1").SetValue(value);
 })
 
})();
```

Создание параграфа

Создание параграфа в документе

**JavaScript Р7**

```
(function()
{
let oDocument = Api.GetDocument(); //Подключаемся к документу
let oParagraph = Api.CreateParagraph(); //Создаем параграф
for(let i = 0; i< 100; i++){ //Цикл
oParagraph.AddText(`${i}`); //Добавляем с помощью AddText форматированною строку в параграф
oDocument.Push(oParagraph); //Добавляем параграф на страницу
}
})();
```

Цвета фона диапазона

Изменение цвета фона выбранного диапазона

**JavaScript Р7**

```
(function()
{
 Api.GetActiveSheet().GetRange("B3").SetFillColor(Api.CreateColorFromRGB(0, 0, 250));
 
})();
```

Жирность для диапазона

Установка жирности для выбранного диапазона

**JavaScript Р7**

```
(function()
{
Api.GetActiveSheet().GetRange("A2").SetBold(true);
})();
```

Цвет шрифта диапазона

Изменение цвета шрифта выбранного диапазона

**JavaScript Р7**

```
(function()
{
Api.GetActiveSheet().GetRange("B4").SetFontColor(Api.CreateColorFromRGB(255, 0, 0));
})();
```

Слияние диапазона ячеек

Слияние выбранного диапазона ячеек

**JavaScript Р7**

```
(function()
{
Api.GetActiveSheet().GetRange("A1:B3").Merge(true);
})();
```

Выделение диапазона для сортировки

Выделение диапазона ячеек для установления сортировки

**JavaScript Р7**

```
(function()
{
 Api.GetActiveSheet().FormatAsTable("A5:AU5");
 
})();
```

Разъединение диапазона

Разъединение выбранного диапазона ячеек

**JavaScript Р7**

```
(function()
{
Api.GetActiveSheet().GetRange("C3:D10").UnMerge();
})();
```

Номера строк из диапазона

Получение номеров строк из выделенной области

**JavaScript Р7**

```
let selection = Api.GetActiveSheet().Selection; // Помещаем выбранный фрагмент в переменную
let rowsArr = []; // Создаём массив строк
let startRow = selection.GetRow(); // Получаем стартовую строку

// Получаем номер строки каждого выделенного элемента
selection.ForEach((x) => {
 endRow = x.GetRow(); 
});

//Получаем массив с номерами всех выбранных строк
for (startRow; startRow < endRow + 2; startRow++) {
 rowsArr.push(startRow);
}
```

Ширина столбца

Установка ширины всего столбца

**JavaScript Р7**

```
(function()
{
Api.GetActiveSheet().SetColumnWidth(1, 25); //Индекс столбца, ширина
})();
```

Очистка диапазона

Очистка ячеек в диапазоне

**JavaScript Р7**

```
let area = secondWorksheet.GetRange("A21:F150"); // Сброс предыдущих значений

function aroundWhiteBorder(el) {
el.SetBorders("Bottom", "Medium", Api.CreateColorFromRGB(255, 255, 255));
el.SetBorders("Top", "Medium", Api.CreateColorFromRGB(255, 255, 255));
el.SetBorders("Left", "Medium", Api.CreateColorFromRGB(255, 255, 255));
el.SetBorders("Right", "Medium", Api.CreateColorFromRGB(255, 255, 255));
}

area.ForEach((x) => {
x.SetValue("");
aroundWhiteBorder(x);
});
```

Шрифт всей страницы

Задание нужного шрифта для всех элементов страницы

**JavaScript Р7**

```
(function()
{
 let oDoc = Api.GetDocument(); // Получаем документ
 let elcount = oDoc.GetElementsCount(); // Получаем количество элементов в документе
 for(let i = 0; i < elcount; i++){ // Перебираем циклом все элементы
 let e = oDoc.GetElement(i); 
 let oTextPr = el.GetTextPr();
 oTextPr.SetFontFamily("Comic Sans MS"); // Прописываем название нужного шрифта
 e.SetTextPr(oTextPr);// Устанавливаем нужный шрифт
 }
})();
```

Список ключей localStorage

Получение списка ключей localStorage в консоли и на активном листе книги

**JavaScript Р7**

```
(function()   //Получение списка ключей в консоль и на лист
{
    let ow = Api.GetActiveSheet();
    console.clear();
    ow.GetRange("A:B").Clear();
    for ( var i = 0, len = localStorage.length; i < len; ++i ) {
       let tkey=localStorage.key(i);
       let titem=localStorage.getItem(tkey);
       console.log("i =",i," key=", tkey,"  item=",titem);
       ow.GetRangeByNumber(i,0).SetValue(tkey);
       ow.GetRangeByNumber(i,1).SetValue(titem);
    }
})();
```

Удаление выбранного ключа localStorage

Удаление выбранного ключа из localStorage. Ключ задается в тексте макроса

**JavaScript Р7**

```
(function()  //Удаление ненужного ключа в LocalStorage
{
    console.clear();
    let mykey='xldb-xmla-connections';
        if (localStorage.getItem(mykey) !== null) {
        localStorage.removeItem(mykey);
        console.log("Ключ ",mykey," удален");
    } else {
        console.log("Ключ ",mykey," не найден");
    }

})();
```

Установка ключа в LocalStorage

Установка ключа A1 из B1 (активный лист) в LocalStorage

**JavaScript Р7**

```
(function()  //Установка ключа A1 из B1 в LocalStorage
{
    let ow = Api.GetActiveSheet();
    let mykey=ow.GetRange("A1").GetValue2();
    let myitem=ow.GetRange("B1").GetValue2();
    localStorage.setItem(mykey,myitem);

})();
```




---
