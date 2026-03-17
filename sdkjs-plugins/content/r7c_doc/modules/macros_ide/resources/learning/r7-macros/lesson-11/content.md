# Справочно: Функции VBA и их аналоги в JavaScript


Математические

Строковые

Массивы

Типы данных

Файлы

Системные

Цвет

Системы счисления

Объекты

Форматирование

Финансы

Дата и время

Указатели

Данные

**Математические функции VBA и**   
**их аналоги в JavaScript (14)**

Математические функции являются основой программы на любом языке программирования. Список всех математических функций на VBA и их аналогов в JavaScript (встроенная библиотека Math) приведен ниже:

VBA: Abs - JavaScript: Math.abs(x).  
VBA: Atn - JavaScript: Math.atan(x)  
VBA: Cos - JavaScript: Math.cos(x)  
VBA: Exp - JavaScript: Math.exp(x)  
VBA: Fix - JavaScript: x < 0 ? Math.ceil(x) : Math.floor(x)  
VBA: Int - JavaScript: Math.floor(x)  
VBA: IsNumeric - JavaScript: !isNaN(x);   
VBA: Log - JavaScript: Math.log(x)  
VBA: Round - JavaScript: Math.round(x)  
VBA: Rnd - JavaScript: Math.random()  
VBA: Sin - JavaScript: Math.sin(x)  
VBA: Sqr - JavaScript: Math.sqrt(x)   
VBA: Sgn - JavaScript: Math.sign(x)  
VBA: Tan - JavaScript: Math.tan(x)

**Функции VBA обработки строк и их аналоги (27)**

При программировании интерфейсов программ на VBA большое значение имеют функции обработки строк. В JavaScript функции, аналогичные функциям обработки строк в VBA имеются в неполном объеме, но создание их аналогов не слишком сложная задача. Приведем список таких функций в VBA и их замену в JavaScript.

ASC

CHR

FILTER

INSTR

InStrRev

Join

LCase

Left

Len

LTrim

Mid

Partition

Replace

Right

RTrim

Space

Spc

Str

StrComp

StrConv

String

StrReverse

Tab

Trim

TypeNam

UCase

Val

Asc(String)=>Integer ASCII AscB(String) =>Byte AscW(String)=>Unicode

Аналог:

```
function asc(str) {    return str.charCodeAt(0); }
```

Chr(CharCode)=>String

Аналог:

```
function chr(asciiCode) {    return String.fromCharCode(asciiCode);}
```

**Filter(SourceArray, Match[, Inclule, [Compare]])=>OutStringArray**

Аналог:

```
function filterArray(sourceArray, match, include = true, caseInsensitive = false) {
    return sourceArray.filter(item => {
        if (typeof item !== 'string') {
            return false;
        }
        // Приведение строки и критерия к нужному регистру, если нужно игнорировать регистр
        const checkItem = caseInsensitive ? item.toLowerCase() : item;
        const checkMatch = caseInsensitive ? match.toLowerCase() : match;
        // Проверка на включение или исключение элементов
        const containsMatch = checkItem.includes(checkMatch);
        return include ? containsMatch : !containsMatch;
    });
}
```

InStr([Start,]String1,String2[,Compare])=>Long

Аналог:

```
function instr(start = 1, string1, string2, caseInsensitive = false) {
    // Приведение строк к нужному регистру, если нужно игнорировать регистр
    const checkString = caseInsensitive ? string1.toLowerCase() : string1;
    const searchString = caseInsensitive ? string2.toLowerCase() : string2;    
    // Поиск позиции первого вхождения
    const position = checkString.indexOf(searchString, start - 1);
    // Возвращаем позицию в 1-индексации или 0, если совпадение не найдено
    return position !== -1 ? position + 1 : 0;
}
```

**InStrRev(StringCheck,StringMatch[,Start[,Compare]])=>Long**

Аналог:

```
function instrRev(stringCheck, stringMatch, start = stringCheck.length, caseInsensitive = false) {
    // Приведение строк к нужному регистру, если нужно игнорировать регистр
    const checkString = caseInsensitive ? stringCheck.toLowerCase() : stringCheck;
    const matchString = caseInsensitive ? stringMatch.toLowerCase() : stringMatch;    
    // Обрезаем строку до позиции start
    const substring = checkString.slice(0, start);    
    // Находим последнюю позицию вхождения
    const position = substring.lastIndexOf(matchString);
    // Возвращаем позицию, если она больше 0, иначе возвращаем -1
    return position !== -1 ? position + 1 : -1;
}
```

Join(SourceArray,[Delimiter])=>String

Аналог:

```
function join(arr, separator = ',') {    return arr.join(separator); }
```

LСase(String)=>String

Аналог:

```
function lcase(str) {    return str.toLowerCase(); }
```

Left(String, Length)=>String

Аналог:

```
function left(str, length) {    return str.slice(0, length); }
```

Len(String)=>Number

Аналог:

```
function len(str) {    return str.length; }
```

LTrim(String)=>String

Аналог:

```
function ltrim(str) {    return str.trimStart();  }
```

Mid(String,start-Long,[length-Long])=>String

Аналог:

```
function mid(str, start, length) { // JavaScript строки индекс от 0, поэтому уменьшаем start на 1 
  // return str.substr(start - 1, length);
return str.substring(start - 1, start - 1 + length); }
```

Partition(Number, Start, Stop, Interval)=>String

Используется редко для статистики.

Replace(Expression,Find,Replace,[Start],[Count],[Compare])=>String

Аналог:

```
function replaceAll(str, search, replacement) {
    // Создаем регулярное выражение для поиска всех вхождений
    // Используем глобальный флаг 'g' для замены всех вхождений
    const regex = new RegExp(search.replace(/[-[\]/{}()*+?.\\^$|]/g, '\\$&'), 'g');
    return str.replace(regex, replacement);
}
```

Right(String)=>String

Аналог:

```
function right(str, length) {
    if (length >= str.length) {
        return str;
    }
    return str.substring(str.length - length);
}
```

RTrim(String)=>String

Аналог:

```
function rtrim(str) {    return str.replace(/\s+$/, ''); 
```

Space(n)=>String

Аналог:

```
function space(n) {    return ' '.repeat(n); }
```

Spc(n)=>String

Аналог:

```
function spc(n) {    return ' '.repeat(n); }
```

Str(Number)=>String

Аналог:

```
function str(value) { return String(value); }
```

StrComp(String1, String2[, Compare])=>Integer

Аналог:

```
function strComp(string1, string2, compareMethod = 0) {
    // если compareMethod == 1, регистр не важен
    if (compareMethod === 1) {
        string1 = string1.toLowerCase();
        string2 = string2.toLowerCase();
    }    
    if (string1 < string2) {
        return -1;
    } else if (string1 > string2) {
        return 1;
    } else {
        return 0;
    }
}
```

StrConv(String,Conversion,[LocaleID])

Специфическая функция, включающая в себя различные варианты простых преобразований (Верхний регистр, нижний, Unicode и т.д.), большинство которых имеют аналоги в JavaScript и разобраны нами в этом разделе.

String(Long, char)=>String

Аналог:

```
function repeatString(count, character) { return character.repeat(count); }
```

StrReverse(String)=>String

Аналог:

```
function strReverse(str) { return str.split('').reverse().join(''); }
```

Tab(n)

Редко используется. Можно использовать в строке вывода "\t".

**Trim(String)=>String**

Аналог:

```
function trim(str) { return str.trim(); }
```

TypeName(Object)=>String

Аналог:

```
function typeName(variable) {
    if (variable === null) return "Null";
    if (variable === undefined) return "Undefined";
    const type = typeof variable;
    switch (type) {
        case "boolean":
            return "Boolean";
        case "number":
            return Number.isInteger(variable) ? "Integer" : "Double";
        case "string":
            return "String";
        case "function":
            return "Function";
        case "object":
            if (Array.isArray(variable)) return "Array";
            if (variable instanceof Date) return "Date";
            return "Object";
        default:
            return "Unknown";
    }
}
```

Ucase(String)=>String

Аналог:

```
function ucase(str) {    return str.toUpperCase(); }
```

Val(String)=>Numeric

Аналог:

```
function val(str) {   // return parseInt(str);
    return parseFloat(str); }

```

Функции работы с массивами в VBA и   
их аналоги в JavaScript (4)

При преобразовании кода VBA, использующего многомерные массивы, возникают проблемы, требующие усложнения логики работы программы. При использовании одномерных массивов следует обратить внимание на настройки VBA (Option Base).

Array

IsArray

LBound

UBound

Array(ParamArray)=>Array

Аналог:

```
let array1 = [1, 2, 3]; // Массив с элементами 1, 2, 3
let array2 = new Array(1, 2, 3); // Массив с элементами 1, 2, 3, созданный через конструктор
//Вариант
function createArray(...elements) {    return elements;  }
let myArray = createArray(1, 2, 3, 4, 5); // Создаст массив [1, 2, 3, 4, 5]

```

IsArray(VarName)=>Boolean

Аналог:

```
function isArray(variable) {    return Array.isArray(variable);  }
```

LBound(ArrayName[,Dimension])=>Long

Аналог:

```
function LBound(arr, dimension) {
    if (!Array.isArray(arr)) {
        throw new Error("Не массив");
    }
    // Проверка на измерение (в JavaScript массивы одномерные, поэтому проверка на измерение)
    if (dimension && dimension !== 1) {
        throw new Error("в JavaScript только одномерные массивы ");
    }    
    // Возвращаемый начальный индекс для массива в JavaScript всегда 0
    return 0;
}

```

UBound(ArrayName[,Dimension])=>Long

Аналог:

```
function UBound(arr, dimension) {
    if (!Array.isArray(arr)) {
        throw new Error("Не массив!");
    }
    // Проверка на измерение (в JavaScript массивы одномерные, поэтому проверка на измерение)
    if (dimension && dimension !== 1) {
        throw new Error("в JavaScript только одномерные массивы");
    }
    // Возвращаем последний индекс массива
    return arr.length - 1;
}

```

Функции преобразования типа данных в VBA и   
их аналоги в JavaScript (11)

CByte

CCur

CDate

CVDate

CDec

CDbl

CInt

CLng

CStr

CSng

CVar

CByte(Expression)=>Byte

Аналог: нет такого типа (0..255) в JavaScript. Можно попробовать так.

```
function CByte(expression) {
    let number = Number(expression);
    if (isNaN(number) || number < 0 || number > 255 || !Number.isInteger(number)) {
        throw new Error("Invalid byte value");
    }
    return number;
}
```

CCur(Expression)=>Currence

Аналог: нет такого типа в JavaScript. Можно попробовать функцию Number с округлением до второго знака после запятой.

```
function CCur(expression) {
    let number = Number(expression);
    if (isNaN(number)) {
        throw new Error("Invalid currency value");
    }
    return number.toFixed(2); // Округляет до двух знаков после запятой для валюты
}
```

CDate(String)=>Date

Аналог (с проверкой корректности формата строки как даты)

```
function CDate(dateString) {
    const date = new Date(dateString);
    // Проверка на корректность даты
    if (isNaN(date)) {
        throw new Error("Invalid date string");
    }
    return date;
} 
```

CVDate(String)=>Date

Устаревшая функция. Есть более современная CDate (Выше).

CDec(Expression)=>Decimal   
(максимальная точность VBA)

Аналог (с проверкой превышения точности для JavaScript):

```
function CDate(dateString) {
    const date = new Date(dateString);
    // Проверка на корректность даты
    if (isNaN(date)) {
        throw new Error("Invalid date string");
    }
    return date;
} 
```

CDbl(Expression)=>Double

Аналог(с проверкой превышения точности дляJavaScript):

```
function CDbl(expression) {
    let number = Number(expression);
    if (isNaN(number)) {
        throw new Error("Invalid number");
    }
    return number;
}
```

CInt(Expression)=>Integer

Аналог:

```
function CInt(expression) {    return Math.round(Number(expression));}
```

CLng(Expression)=>Long

Аналог:

```
function CLng(expression) {    return parseInt(expression, 10); }
```

CStr(Expression)=>String

Аналог:

```
function CStr(expression) {    return String(expression); }
```

CSng(expression)=>Float

Аналог:

```
function CSng(expression) {    return parseFloat(expression);} 
```

CVar(Expression)=>Variant

Аналог: Нет нужды в аналоге, так как используется динамическая типизация. Однако имеются и функции явной типизации:

```
function toString(value) {    return String(value); }
function toNumber(value) {    return Number(value); }
function toBoolean(value) {    return Boolean(value); }
function toObject(value) {    return Object(value); }

```

Функции работы с файлами (12)

Из-за ограничений JavaScript в Р7, разобранных выше, аналогов нет. Возможно использование аналогов в Node.js или файловых потоках.  
CurDir[ (Drive) ]=>Variant(String)CurDir$=>String  
Dir [(PathName[, Attributes])]=>String  
EOF(FileNumber)=>Boolean (Integer)  
FreeFile([RangeNumber])=>Integer (0-511)  
FileLen(PathName)=>Long  
FileDateTime(PathName)=>Variant(Date)  
FileAttr(FileNumber[,ReturnType])=>Long  
GetAttr(Pathname)=>Integer  
Input(Number, [#]FileNumber)=>String  
LOF(FileNumber)=>Long  
Loc(FileNumber)=>Long  
Seek(FileNumber)=>Long

Системные функции (16)

Switch

Shell

IsNull

IsMissing

IsError

IsEmpty

IMEStatus

GetSetting

GetAllSettings

Error

Erl

Environ

DoEvents

CVErr

Command

CallByName

Switch(Expr-1, Value-1[, Expr-2, Value-2 … [, Expr-n,Value-n]])

Аналог:

```
function switchCase(expr, ...cases) {
    for (let i = 0; i < cases.length; i += 2) {
        const caseExpr = cases[i];
        const caseValue = cases[i + 1];
        if (expr === caseExpr) {
            return caseValue;
        }
    }
    // Если не найдено совпадений, возвращаем значение по умолчанию
    return undefined;
}
```

Shell(PathName,[WindowStyle])=>Variant(Double)

Аналог: В связи с ограничениями, разобранными выше, аналога нет.  
Для общего развития в Node.js:

```
const { exec } = require('child_process');
function shell(command, windowStyle) {
    // Список возможных значений для windowStyle
    const windowStyles = {
        'hide': 0,       // Скрытое окно
        'normal': 1,     // Обычное окно
        'minimized': 2,  // Свернутое окно
        'maximized': 3   // Развернутое окно
    };
    // Устанавливаем параметры окна, если указано
    let style = windowStyles[windowStyle] !== undefined ? windowStyles[windowStyle] : windowStyles['normal'];
    // Запускаем команду
    exec(command, (error, stdout, stderr) => {
        if (error) {
            console.error(`Ошибка: ${error.message}`);
            return;
        }
        if (stderr) {
            console.error(`Стандартный поток ошибок: ${stderr}`);
            return;
        }
        console.log(`Стандартный вывод: ${stdout}`);
    });
}
// Пример использования
shell('notepad', 'normal');  // Запускает Notepad с обычным окном
```

IsNull(Expression=>Boolean

Аналог:

```
function isNull(expression) {    return expression === null || expression === undefined; }
```

IsMissing(ArgName)=>Boolean

Аналог:

```
function isMissing(arg) {    return typeof arg === 'undefined';}
```

IsError(Expression)=>Boolean

Аналог:

```
function isError(expression) {    return expression instanceof Error;}
```

IsEmpty(Expression)=>Boolean

Аналог:

```
function isEmpty(expression) {    return expression === undefined || expression === null || expression === "";}
```

IMEStatus[( )]=>Integer

Аналог не нужен. Только для восточно-азиатских версий.

GetSetting(AppName,Section,Key[,Default])=>String

Специфичная функция для VBA (параметр реестра для VB и VBA). Доступа к реестру в JavaScript нет. Можно попробовать использовать локальное хранилище localStorage:

```
function getSetting(appName, section, key, defaultValue) {
    const storageKey = `${appName}.${section}.${key}`;
    const value = localStorage.getItem(storageKey);
    return value !== null ? value : defaultValue;
}
// Примеры использования
localStorage.setItem('MyApp.Settings.Theme', 'dark');
console.log(getSetting('MyApp', 'Settings', 'Theme', 'light')); // "dark"
```

GetAllSettings(Appname,Section)=>String

Специфичная функция для VBA (параметр реестра для VB и VBA). Доступа к реестру в JavaScript нет. Можно попробовать использовать локальное хранилище localStorage:

```
function setSetting(appName, section, key, value) {
    const storageKey = `${appName}.${section}.${key}`;
    localStorage.setItem(storageKey, value);
}
function getAllSettings(appName, section) {
    const settings = {};
    const prefix = `${appName}.${section}.`;
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key.startsWith(prefix)) {
            const settingKey = key.substring(prefix.length);
            settings[settingKey] = localStorage.getItem(key);
        }
    }
    return settings;
}
// Установка значений
setSetting('MyApp', 'Settings', 'Theme', 'dark');
setSetting('MyApp', 'Settings', 'FontSize', '16px');
setSetting('MyApp', 'User', 'Name', 'JohnDoe');
// Получение всех настроек
console.log(getAllSettings('MyApp', 'Settings'));  // { Theme: 'dark', FontSize: '16px' }
console.log(getAllSettings('MyApp', 'User'));  // { Name: 'JohnDoe' }
```

Error[(ErrorNumber)]=>Variant(String)

Аналог:

```
// Определяем свой набор сообщений об ошибках
const errorMessages = {
    1: "Ошибка: Недопустимый аргумент.",
    2: "Ошибка: Доступ запрещен.",
    3: "Ошибка: Элемент не найден.",
    4: "Ошибка: Превышен лимит времени.",
    5: "Ошибка: Недостаточно памяти.",
};
function getError(errorNumber) {
    return errorMessages[errorNumber] || `Неизвестная ошибка: ${errorNumber}`;
}
```

Erl()=Long

Аналог: Аналог может не сработать, требует проверки!:

```
function getErrorLine(error) {
    if (error.stack) {
        // Разбираем стек вызовов
        const stackLines = error.stack.split('\n');
        // Возвращаем первую строку стека, которая содержит номер строки ошибки
        // Обычно это третья строка стека, но это может варьироваться
        const lineInfo = stackLines[1].match(/:(\d+):\d+\)?$/);
        if (lineInfo) {
            return parseInt(lineInfo[1], 10);
        }
    }
    return null;
}
// Пример использования
try {
    // Код, который может вызвать ошибку
    nonExistentFunction(); // Эта строка вызовет ошибку
} catch (error) {
    const line = getErrorLine(error);
    console.log(`Ошибка произошла на строке: ${line}`);
}
```

Environ(Expression)=>String

Специфичная функция для VBA. Аналога нет.  
Для общего развития в Node.js:

```
function getEnvironmentVariable(name) {    return process.env[name] || null;}
```

DoEvents()=>Integer

Аналог:

```
function doEvents(callback) {
    setTimeout(callback, 0);
}
// Пример использования
console.log("Start");
doEvents(() => {
    console.log("DoEvents callback");
});
console.log("End");
```

CVErr(errornumber)

Аналог (имя функции для лучшего понимания функционирования не совпадает с именем исходной функции):

```
function createCustomError(errorNumber) {
    let error = new Error(`Custom error with number: ${errorNumber}`);
    error.number = errorNumber;
    return error;
}
// Примеры использования
try {
    throw createCustomError(1001);
} catch (e) {
    console.log(e.message); // "Custom error with number: 1001"
    console.log(e.number);  // 1001
}
```

Command()=>String

Аналога нет, так как нет прямого доступа к командной строке.  
Для общего развития в Node.js:

```
// Получение аргументов командной строки
const args = process.argv.slice(2); // slice(2) чтобы исключить первые два элемента (путь к Node.js и скрипту)
// Вывод аргументов
console.log("Command-line arguments:");
args.forEach((arg, index) => {    console.log(`Argument ${index + 1}: ${arg}`); });
```

CallByName(Object,ProcName,CallType,[Args() ])

Аналог:

```
function callByName(object, methodName, args) {
    if (typeof object[methodName] === 'function') {
        // Если метод существует, вызываем его с переданными аргументами
        return object[methodName](...args);
    } else {
        throw new Error(`Method ${methodName} does not exist on the object.`);
    }
}
// Пример объекта с методами
const myObject = {
    greet(name) {
        return `Hello, ${name}!`;
    },
    add(a, b) {
        return a + b;
    }
};
// Использование функции callByName
try {
    const greeting = callByName(myObject, 'greet', ['Alice']);
    console.log(greeting); // Output: Hello, Alice!

    const sum = callByName(myObject, 'add', [5, 3]);
    console.log(sum); // Output: 8
} catch (error) {
    console.error(error.message);
}
```

Функции обработки цвета (2)

RGB

QBColor

RGB(Red, Green, Blue)=>Long

Аналог для CSS (возвращает String ‘rgb(r,g,b)’):

```
function RGB(red, green, blue) {
    // Проверка, чтобы значения были в пределах от 0 до 255
    if (red < 0 || red > 255 || green < 0 || green > 255 || blue < 0 || blue > 255) {
        throw new Error('Values must be between 0 and 255');
    }
    // Преобразование в строку формата RGB
    return `rgb(${red}, ${green}, ${blue})`;
}
// Примеры использования
try {
    const color1 = RGB(255, 0, 0); // Красный цвет
    console.log(color1); // Output: rgb(255, 0, 0)
    const color4 = RGB(128, 128, 128); // Серый цвет
    console.log(color4); // Output: rgb(128, 128, 128)
} catch (error) {
    console.error(error.message);
}
```

QBColor(Color)=>Long

Аналог:

```
function QBColor(index) {
    // Палитра цветов QB64
    const colors = [
        "#000000", // 0: Black
        "#FF0000", // 1: Red
        "#00FF00", // 2: Green
        "#FFFF00", // 3: Yellow
        "#0000FF", // 4: Blue
        "#FF00FF", // 5: Magenta
        "#00FFFF", // 6: Cyan
        "#C0C0C0", // 7: Light Gray
        "#808080", // 8: Gray
        "#FF0000", // 9: Red (same as index 1)
        "#00FF00", // 10: Green (same as index 2)
        "#FFFF00", // 11: Yellow (same as index 3)
        "#0000FF", // 12: Blue (same as index 4)
        "#FF00FF", // 13: Magenta (same as index 5)
        "#00FFFF", // 14: Cyan (same as index 6)
        "#FFFFFF"  // 15: White
    ];
    // Проверка, чтобы индекс был в допустимом диапазоне
    if (index < 0 || index > 15) {
        throw new Error('Index must be between 0 and 15');
    }
    return colors[index];
}
// Примеры использования
try {
    console.log(QBColor(0));  // Output: #000000 (Black)
    console.log(QBColor(3));  // Output: #FFFF00 (Yellow)
    console.log(QBColor(7));  // Output: #C0C0C0 (Light Gray)
    console.log(QBColor(15)); // Output: #FFFFFF (White)
} catch (error) {
    console.error(error.message);
}
```

Функции преобразования чисел в   
разные системы счисления (3)

VarType (VarName)=>Integer

Аналог:

```
function VarType(value) {
    if (value === null) return 'Null';
    if (value === undefined) return 'Undefined';
    if (typeof value === 'boolean') return 'Boolean';
    if (typeof value === 'number') return 'Number';
    if (typeof value === 'string') return 'String';
    if (typeof value === 'object') {
        if (Array.isArray(value)) return 'Array';
        if (value instanceof Date) return 'Date';
        return 'Object';
    }
    if (typeof value === 'function') return 'Function';
    return 'Unknown';
}
```

Oct(Number)=>Variant(String)

Аналог:

```
function Oct(number) {
    if (typeof number !== 'number' || isNaN(number) || !Number.isInteger(number)) {
        throw new TypeError('Input must be an integer.');
    }
    return number.toString(8);
}
```

Hex(Number)=>Variant(String)

Аналог:

```
function Hex(number) {
    if (typeof number !== 'number' || isNaN(number) || !Number.isInteger(number)) {
        throw new TypeError('Input must be an integer.');
    }
    return number.toString(16).toUpperCase();
}
```

Функции работы с объектами (4)

IsObject(Expression)=>Boolean

Аналог:

```
function IsObject(expression) {    return typeof expression === 'object' && expression !== null;}
```

GetObject([Pathname] [Class])

Аналог: Прямого аналога в JavaScript нет в связи с ограничениями, разобранными выше. Возможна частичная аналогия в частных случаях.

1.Непосредственная работа с объектами в JavaScript.

```
const myObject = {    //Объект
    name: 'John',
    age: 30,
    greet: function() {
        console.log('Hello, ' + this.name);
    }
};
// Доступ к свойствам объекта
console.log(myObject.name); // John
myObject.greet(); // Hello, John
```

2.Объекты в DOM.

```
const myElement = document.getElementById('myElementId');
console.log(myElement);
```

3.Объекты API через fetch()

```
fetch('https://api.example.com/data')
    .then(response => response.json())
    .then(data => {
        console.log('Data from API:', data);
    })
    .catch(error => {
        console.error('Error fetching data:', error);
    });
```

4.Node.js

```
const fs = require('fs');
// Чтение содержимого файла
fs.readFile('path/to/file.txt', 'utf8', (err, data) => {
    if (err) {
        console.error('Error reading file:', err);
        return;
    }
    console.log('File content:', data);
});
```

GetAutoServerSettings([Progid], [CLSID])=>Variant

Аналог: Прямого аналога в JavaScript нет в связи с ограничениями, разобранными выше.

CreateObject(Class,[ServerName])=>Object

Аналог: Прямого аналога в JavaScript нет в связи с ограничениями, разобранными выше.  
Для общего развития в Node.js:

```
const fs = require('fs'); // Built-in Node.js module for file system operations
// Create an instance of a class from a Node.js module
fs.readFile('example.txt', 'utf8', (err, data) => {
    if (err) throw err;
    console.log(data);
});
```

Функции форматирования (5)

FormatPercent(Expression[,NumDigitsAfterDecimal [,IncludeLeadingDigit [,UseParensForNegativeNumbers [,GroupDigits]]]])=>String or Number

Аналог:

```
function formatPercent(value, numDigitsAfterDecimal = 2, includeLeadingDigit = true, useParensForNegativeNumbers = false, groupDigits = false) {
    // Convert value to a percentage
    let percentage = (value * 100).toFixed(numDigitsAfterDecimal);
    // Handle leading zero
    if (!includeLeadingDigit && percentage.startsWith('0.')) {
        percentage = percentage.slice(1); // Remove the leading zero
    }
    // Handle grouping of digits
    if (groupDigits) {
        let parts = percentage.split('.');
        parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ','); // Add commas for thousands
        percentage = parts.join('.');
    }
    // Handle parentheses for negative numbers
    if (useParensForNegativeNumbers && value < 0) {
        percentage = `(${percentage})`;
    } else if (value < 0) {
        percentage = `-${percentage}`;
    }
    // Append percentage sign
    return `${percentage}%`;
}
// Примеры использования
console.log(formatPercent(0.12345)); // "12.35%"
console.log(formatPercent(0.12345, 1)); // "12.3%"
console.log(formatPercent(0.12345, 2, false)); // "12.35%"
console.log(formatPercent(-0.12345, 2, true, true)); // "(12.35%)"
console.log(formatPercent(12345.678, 2, true, false, true)); // "1,234,567.68%"
```

FormatNumber(Expression[,NumDigitsAfterDecimal [,IncludeLeadingDigit [,UseParensForNegativeNumbers [,GroupDigits]]]])=>Variant(String)

Аналог:

```
function formatNumber(value, numDigitsAfterDecimal = 2, includeLeadingDigit = true, useParensForNegativeNumbers = false, groupDigits = false) {
    let formattedNumber = value.toFixed(numDigitsAfterDecimal);
    // Обработка ведущего нуля
    if (!includeLeadingDigit && formattedNumber.startsWith('0.')) {
        formattedNumber = formattedNumber.slice(1); // Удаляем ведущий ноль
    }
    // Обработка группировки цифр (тысячные разделители)
    if (groupDigits) {
        let parts = formattedNumber.split('.'); // Разделяем целую и дробную части
        parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ','); // Добавляем запятые для тысячных
        formattedNumber = parts.join('.'); // Соединяем части обратно
    }
    // Обработка скобок для отрицательных чисел
    if (useParensForNegativeNumbers && value < 0) {
        formattedNumber = `(${formattedNumber})`; // Оборачиваем отрицательное число в скобки
    } else if (value < 0) {
        formattedNumber = `-${formattedNumber}`; // Добавляем минус перед числом
    }
    return formattedNumber;
}
// Примеры использования
console.log(formatNumber(12345.6789, 1)); // "12,345.7" - 1 знак после запятой
console.log(formatNumber(12345.6789, 2, false)); // "12,345.68" - ведущий ноль не включается
console.log(formatNumber(-12345.6789, 2, true, true)); // "(12,345.68)" - отрицательное число в скобках
console.log(formatNumber(12345.6789, 2, true, false, true)); // "12,345.68" - группировка цифр включена
```

FormatDateTime(Date[,NamedFormat])=>Variant(String)

Аналог:

```
function formatDateTime(date, namedFormat) {
    if (!(date instanceof Date)) {
        throw new TypeError("Первый аргумент должен быть объектом Date.");
    }
    // Если формат не указан, используем формат по умолчанию (краткий формат даты и времени)
    namedFormat = namedFormat || "short";
    // Функция для добавления нулей в однозначные числа
    const pad = (num) => num.toString().padStart(2, '0');
    // Определяем форматирование в зависимости от указанного namedFormat
    switch (namedFormat.toLowerCase()) {
        case "short":
            // Краткий формат даты и времени
            return `${pad(date.getDate())}/${pad(date.getMonth() + 1)}/${date.getFullYear()}
                            ${pad(date.getHours())}:${pad(date.getMinutes())}`;
        case "long":
            // Длинный формат даты и времени
            return `${date.getDate()} ${date.toLocaleString('default', { month: 'long' })} ${date.getFullYear()}, 
                          ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
        case "shortdate":
            // Краткий формат даты
            return `${pad(date.getDate())}/${pad(date.getMonth() + 1)}/${date.getFullYear()}`;
        case "longdate":
            // Длинный формат даты
            return `${date.getDate()} ${date.toLocaleString('default', { month: 'long' })} ${date.getFullYear()}`;
        case "shorttime":
            // Краткий формат времени
            return `${pad(date.getHours())}:${pad(date.getMinutes())}`;
        case "longtime":
            // Длинный формат времени
            return `${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
        default:
            throw new Error("Неизвестный формат даты/времени.");
    }
}
// Примеры использования
const now = new Date();
console.log(formatDateTime(now)); // Краткий формат даты и времени по умолчанию
console.log(formatDateTime(now, "short")); // Краткий формат даты и времени
console.log(formatDateTime(now, "long")); // Длинный формат даты и времени
console.log(formatDateTime(now, "shortdate")); // Краткий формат даты
console.log(formatDateTime(now, "longdate")); // Длинный формат даты
console.log(formatDateTime(now, "shorttime")); // Краткий формат времени
console.log(formatDateTime(now, "longtime")); // Длинный формат времени
```

FormatCurrency(Expression[,NumDigitsAfterDecimal   
 [,IncludeLeadingDigit,[UseParensForNegativeNumbers[,GroupDigits]]]])=>  
 Variant(String)

Аналог:

```
function formatCurrency(expression, numDigitsAfterDecimal = 2, includeLeadingDigit = false,
                              useParensForNegativeNumbers = false, groupDigits = true) {
    if (typeof expression !== 'number' || isNaN(expression)) {
        throw new TypeError("Первый аргумент должен быть числом.");
    }
    // Создаем объект Intl.NumberFormat для форматирования чисел в денежном формате
    const options = {
        style: 'currency',
        currency: 'USD', // Установите валюту по умолчанию, если необходимо
        minimumFractionDigits: numDigitsAfterDecimal,
        maximumFractionDigits: numDigitsAfterDecimal,
        useGrouping: groupDigits
    };
    // Создаем экземпляр Intl.NumberFormat
    const formatter = new Intl.NumberFormat('en-US', options);
    // Форматируем число
    let formattedNumber = formatter.format(expression);
    // Если нужно добавить ведущий ноль, если число меньше 1
    if (includeLeadingDigit && expression < 1 && expression > -1) {
        formattedNumber = formattedNumber.replace(/^(?!\$\s*0)/, '$0');
    }
    // Если нужно использовать скобки для отрицательных чисел
    if (useParensForNegativeNumbers && expression < 0) {
        formattedNumber = formattedNumber.replace(/^\(\$/, '$(').replace(/\)$/, ')');
    }
    return formattedNumber;
}
// Примеры использования
console.log(formatCurrency(1234.567)); // Форматирование с двумя знаками после запятой и группировкой
console.log(formatCurrency(1234.567, 3)); // Форматирование с тремя знаками после запятой
console.log(formatCurrency(1234.567, 2, true)); // Форматирование с ведущим нулем, если число < 1
console.log(formatCurrency(1234.567, 2, false, false, false)); // Форматирование без группировки
```

Format (Expression[, Format[, FirstDayOfWeek[, FirstWeekOfYear]]])=>Variant (String)

Аналог: Единого аналога нет. Для разных типов входных данных:

1.Число.

```
function formatNumber(expression, format) {
    if (typeof expression !== 'number' || isNaN(expression)) {
        throw new TypeError("Первый аргумент должен быть числом.");
    }
    // Применяем форматирование чисел, если формат задан
    if (format === 'currency') {
        return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(expression);
    } else if (format === 'percent') {
        return new Intl.NumberFormat('en-US', { style: 'percent' }).format(expression);
    } else if (format === 'fixed') {
        return expression.toFixed(2); // Форматирование числа с двумя знаками после запятой
    } else {
        return expression.toString(); // По умолчанию просто преобразуем число в строку
    }
}
// Примеры использования
console.log(formatNumber(1234.567, 'currency')); // $1,234.57
console.log(formatNumber(0.1234, 'percent')); // 12.34%
console.log(formatNumber(1234.567, 'fixed')); // 1234.57
```

2.Дата.

```
function formatDate(expression, format, firstDayOfWeek = 0, firstWeekOfYear = 0) {
    if (!(expression instanceof Date)) {
        throw new TypeError("Первый аргумент должен быть объектом Date.");
    }
    const options = {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit'
    };

    if (format === 'shortDate') {
        return new Intl.DateTimeFormat('en-US', options).format(expression);
    } else if (format === 'longDate') {
        return new Intl.DateTimeFormat('en-US', { ...options, weekday: 'long' }).format(expression);
    } else if (format === 'time') {
        return new Intl.DateTimeFormat('en-US', { hour: '2-digit', minute: '2-digit' }).format(expression);
    } else {
        return expression.toLocaleDateString(); // По умолчанию просто преобразуем дату в строку
    }
}
// Примеры использования
console.log(formatDate(new Date(), 'shortDate')); // 07/10/2024
console.log(formatDate(new Date(), 'longDate')); // Wednesday, July 10, 2024
console.log(formatDate(new Date(), 'time')); // 10:30 AM
```

Финансовые функции (13)

Финансовые функции имеют специализованное применение, но для некоторых задач могут облегчить процесс программирования.

SYD(Cost, Salvage, Life, Period)=>Double

Вычисление годовой амортизации фондов для заданного года с использованием метода линейной амортизации.  
Аналог:

```
function SYD(cost, salvage, life, period) {
// life - Полный срок службы актива
// cost - Первоначальная стоимость актива
// salvage - Остаточная стоимость актива в конце срока службы
// period - Период, для которого рассчитывается амортизация
    if (life <= 0 || period <= 0 || period > life) {
        throw new Error("Некорректные значения для срока службы или периода.");
    }
    // Рассчитываем количество лет
    const sumOfYearsDigits = (life * (life + 1)) / 2;
    // Рассчитываем амортизацию для указанного периода
    const depreciation = ((cost - salvage) * (life - period + 1)) / sumOfYearsDigits;
    
    return depreciation;
}
// Примеры использования
console.log(SYD(10000, 2000, 5, 1)); // Амортизация за 1-й период
console.log(SYD(10000, 2000, 5, 2)); // Амортизация за 2-й период
console.log(SYD(10000, 2000, 5, 3)); // Амортизация за 3-й период
```

SLN(Cost, Salvage, Life)=>Double

Вычисление годовой амортизации фондов при фиксированных нормативах амортизации.  
Аналог:

```
function SLN(cost, salvage, life) {
// life - Полный срок службы актива
// cost - Первоначальная стоимость актива
// salvage - Остаточная стоимость актива в конце срока службы
    if (life <= 0) {
        throw new Error("Срок службы должен быть больше нуля.");
    }
    
    // Рассчитываем амортизацию по методу прямолинейной амортизации
    const depreciation = (cost - salvage) / life;
    
    return depreciation;
}
```

Rate(NPer, Pmt, PV, [FV], [Due], [Guess])=>Double

Вычисление расчета выплат по закладной (например за дом), аннуитетов или итогов накоплений при ежемесячных банковских взносах.  
Аналог:

```
function Rate(nper, pmt, pv, fv = 0, due = 0, guess = 0.1) {
// nper - Общее количество периодов
// pmt - Платеж за период
// pv - Приведенная стоимость (сумма займа)
// fv - Будущая стоимость (в конце).
// due - Тайминг платежей: 0 - конец периода, 1 - начало периода.
// guess - Начальное предположение для процентной ставки.
    const epsMax = 1e-10; // Максимальная ошибка
    const iterMax = 50; // Максимальное количество итераций
    let y, y0, y1, x0, x1 = 0, f = 0, i = 0;
    let rate = guess;
    if (Math.abs(rate) < epsMax) {
        y = pv * (1 + nper * rate) + pmt * (1 + rate * due) * nper + fv;
    } else {
        f = Math.exp(nper * Math.log(1 + rate));
        y = pv * f + pmt * (1 / rate + due) * (f - 1) + fv;
    }
    y0 = pv + pmt * nper + fv;
    y1 = pv * f + pmt * (1 / rate + due) * (f - 1) + fv;
    x0 = 0.0;
    x1 = rate;
    i = y0 * y1 >= 0 ? 0 : 1;
    while ((Math.abs(y0 - y1) > epsMax) && (i < iterMax)) {
        rate = (y1 * x0 - y0 * x1) / (y1 - y0);
        x0 = x1;
        x1 = rate;

        if (Math.abs(rate) < epsMax) {
            y = pv * (1 + nper * rate) + pmt * (1 + rate * due) * nper + fv;
        } else {
            f = Math.exp(nper * Math.log(1 + rate));
            y = pv * f + pmt * (1 / rate + due) * (f - 1) + fv;
        }
        y0 = y1;
        y1 = y;
        i++;
    }
    return rate;
}
```

PV(rate, nper, pmt, fv, type)=>Double

Вычисление приведенной стоимости (Present Value) аннуитета.  
Аналог:

```
function PV(rate, nper, pmt, fv = 0, type = 0) {
// nper - Общее количество периодов
// rate - Процентная ставка за период
// pmt - Платеж за период
// type - Тайминг платежей: 0 - конец периода, 1 - начало периода.
// fv - Будущая стоимость (в конце).
    let pv;
    if (rate === 0) {
        pv = -pmt * nper - fv;
    } else {
        pv = (-pmt * (1 + rate * type) * (1 - Math.pow(1 + rate, -nper)) / rate - fv * Math.pow(1 + rate, -nper));
    }
    return pv;
}
```

PPmt(Rate, Per, NPer, PV, [FV], [Due])=>Double

Вычисление платежа по основной сумме кредита (Principal Payment) на конкретный период.  
Аналог:

```
function PPmt(rate, per, nper, pv, fv = 0, type = 0) {
// rate - Процентная ставка за период
// per - Период, для которого рассчитывается платеж
// nper - Общее количество периодов
// pv - Приведенная стоимость (сумма займа)
// fv - Будущая стоимость (в конце).
// type - Тайминг платежей: 0 - конец периода, 1 - начало периода.
    const pmt = Pmt(rate, nper, pv, fv, type);    
    // Рассчитываем процентный платеж на данный период
    let ipmt;
    if (type === 1) {
        if (per === 1) {
            ipmt = 0;
        } else {
            ipmt = (pv * rate) * Math.pow(1 + rate, per - 2);
        }
    } else {
        ipmt = (pv * rate) * Math.pow(1 + rate, per - 1);
    }
    // Возвращаем разницу между общим платежом и процентным платежом
    return pmt - ipmt;
}
//Рассчитывает периодический платеж по аннуитету.
function Pmt(rate, nper, pv, fv = 0, type = 0) {
    let pmt;
    if (rate === 0) {
        pmt = -(pv + fv) / nper;
    } else {
        pmt = - (pv * Math.pow(1 + rate, nper) + fv) * rate / ((1 + rate * type) * (Math.pow(1 + rate, nper) - 1));
    }
    return pmt;
}
// Примеры использования
console.log(PPmt(0.05 / 12, 1, 60, 10000)); // Платеж по основной сумме кредита на первый период
console.log(PPmt(0.05 / 12, 12, 60, 10000)); // Платеж по основной сумме кредита на двенадцатый период
```

NPV(Rate, ValueArray())=>Double

Вычисление приведенного на текущий момент сальдо ряда будущих финансовых операций с учетом уценки капитала по модели финансового потока. Иначе, вычисление чистой приведенной стоимости (Net Present Value) ряда денежных потоков, дисконтированных по заданной ставке.  
Аналог:

```
function NPV(rate, valueArray) {
// rate - Дисконтная ставка за период
// {Array<number>} valueArray - Массив денежных потоков
// {number} - Чистая приведенная стоимость
    if (!Array.isArray(valueArray)) {
        throw new Error("valueArray should be an array");
    }
    let npv = 0;
    for (let i = 0; i < valueArray.length; i++) {
        npv += valueArray[i] / Math.pow(1 + rate, i + 1);
    }
    return npv;
}
// Пример использования
const rate = 0.1; // Дисконтная ставка
const cashFlows = [-1000, 300, 400, 500, 600]; // Массив денежных потоков
console.log(NPV(rate, cashFlows)); // Чистая приведенная стоимость
```

NPer(Rate, Pmt, PV, [FV], [Due])=>Double

Вычисление количества периодов для аннуитета на основе постоянных периодических платежей и постоянной процентной ставки.  
Аналог:

```
function NPer(rate, pmt, pv, fv = 0, due = 0) {
    if (rate === 0) {
        return -(pv + fv) / pmt;
    } else {
        const adjustedRate = 1 + rate * due;
        return Math.log((pmt * adjustedRate - fv * rate) / (pv * rate + pmt * adjustedRate)) / Math.log(1 + rate);
    }
}
// Примеры использования
const rate = 0.05; // Процентная ставка за период
const pmt = -200; // Платеж за период
const pv = 1000; // Приведенная стоимость
const fv = 0; // Будущая стоимость (по умолчанию 0)
const due = 0; // Платежи в конце периода (по умолчанию)
console.log(NPer(rate, pmt, pv, fv, due)); // Количество периодов
```

MIRR(ValueArray(), FinanceRate, ReinvestRate)=>Double

Вычисление модифицированной внутренней нормы прибыли для ряда периодических денежных потоков.  
Аналог:

```
function MIRR(values, financeRate, reinvestRate) {
    let npvPositive = 0;
    let npvNegative = 0;
    for (let i = 0; i < values.length; i++) {
        if (values[i] > 0) {
            npvPositive += values[i] / Math.pow(1 + reinvestRate, i);
        } else {
            npvNegative += values[i] / Math.pow(1 + financeRate, i);
        }
    }
    if (npvNegative === 0 || npvPositive === 0) {
        return 0;
    }
    const totalPeriods = values.length - 1;
    const mirr = 
        Math.pow(-npvPositive * Math.pow(1 + reinvestRate, totalPeriods) / npvNegative, 1 / totalPeriods) - 1;    
    return mirr;
}
// Пример использования
const values = [-1000, 200, 300, 400, 500]; // Пример денежных потоков
const financeRate = 0.05; // Процентная ставка для отрицательных денежных потоков
const reinvestRate = 0.10; // Процентная ставка для положительных денежных потоков
console.log(MIRR(values, financeRate, reinvestRate));
```

IRR(ValueArray()[,Guess])=>Double

Вычисление внутренней нормы прибыли для ряда периодических денежных потоков.  
Аналог:

```
function IRR(values, guess = 0.1) {
    const maxIteration = 1000;
    const precision = 1e-7;
    let rate = guess;
    let iteration = 0;
    let newRate;
    while (iteration < maxIteration) {
        let npv = 0;
        let dNpV = 0;
        for (let i = 0; i < values.length; i++) {
            npv += values[i] / Math.pow(1 + rate, i);
            dNpV -= (i * values[i]) / Math.pow(1 + rate, i + 1);
        }
        newRate = rate - npv / dNpV;
        if (Math.abs(newRate - rate) < precision) {
            return newRate;
        }
        rate = newRate;
        iteration++;
    }
    throw new Error("IRR calculation did not converge");
}
// Пример использования
const values = [-1000, 200, 300, 400, 500]; // Пример денежных потоков (расходы/доходы)
const guess = 0.1; // Начальное предположение для IRR (необязательно)
console.log(IRR(values, guess)); // Результат: внутренняя норма прибыли
```

IPmt(Rate,Per,NPer,PV[,FV[,Due]])=>Double

Вычисление процентного платежа за определенный период инвестиции, основанного на фиксированных периодических платежах и фиксированной процентной ставке.  
Аналог:

```
function IPmt(rate, per, nper, pv, fv = 0, type = 0) {
    if (per < 1 || per > nper) {
        throw new Error("Период должен быть между 1 и общим количеством периодов (nper).");
    }
    // Рассчитываем общую сумму платежа (PMT)
    const PMT = (rate !== 0) 
        ? (pv * rate * Math.pow(1 + rate, nper) + fv * rate) / (Math.pow(1 + rate, nper) - 1)
        : (pv + fv) / nper;
    // Рассчитываем процентную часть платежа
    const IPMT = (rate !== 0)
        ? (pv * Math.pow(1 + rate, per - 1)) * rate - (type === 1 ? PMT / (1 + rate) : 0)
        : pv * rate;
    return IPMT;
}
// Пример использования
const rate = 0.05 / 12; // Месячная процентная ставка
const per = 5; // Номер периода
const nper = 360; // Общее количество периодов
const pv = 100000; // Приведенная стоимость
const fv = 0; // Будущая стоимость (по умолчанию 0)
const type = 0; // Платеж в конце периода (по умолчанию 0)
console.log(IPmt(rate, per, nper, pv, fv, type)); // Результат: процентный платеж
```

FV(Rate, NPer, Pmt[, PV[, Due]])=>Double

Вычисление будущей стоимости инвестиции или сбережений на основе фиксированных периодических платежей и фиксированной процентной ставки.  
Аналог:

```
function FV(rate, nper, pmt, pv = 0, due = 0) {
    if (rate === 0) {
        return -(pv + pmt * nper);
    }
    const typeAdjustment = due === 1 ? 1 : 0;
    const futureValue = -((pv * Math.pow(1 + rate, nper)) + (pmt * (1 + rate * typeAdjustment) * ((Math.pow(1 + rate, nper) - 1) / rate)));
    return futureValue;
}
// Пример использования
const rate = 0.05 / 12; // Месячная процентная ставка
const nper = 360; // Общее количество периодов
const pmt = -1000; // Ежемесячный платеж
const pv = 0; // Приведенная стоимость (по умолчанию 0)
const due = 0; // Платеж в конце периода (по умолчанию 0)
console.log(FV(rate, nper, pmt, pv, due)); // Результат: будущая стоимость
```

DDB(Cost, Salvage, Life, Period, [Factor])=>Double

Вычисление амортизации по методу двойного уменьшающегося остатка за определенный период.  
Аналог:

```
function DDB(cost, salvage, life, period, factor = 2) {
    let bookValue = cost;
    let depreciation = 0;
    for (let i = 1; i <= period; i++) {
        depreciation = Math.min((bookValue * factor) / life, bookValue - salvage);
        bookValue -= depreciation;
    }
    return depreciation;
}
// Пример использования
const cost = 10000; // Начальная стоимость
const salvage = 1000; // Ликвидационная стоимость
const life = 5; // Срок службы (в периодах)
const period = 3; // Период для которого рассчитывается амортизация
const factor = 2; // Коэффициент ускорения (по умолчанию 2)
console.log(DDB(cost, salvage, life, period, factor)); // Результат: амортизация за указанный период
```

Функции работы с датами и временем (21)

Year(Date)=>Variant(Integer)

Аналог:

```
function getYear(date) {    return date.getFullYear();}
```

WeekdayName(Weekday[, Abbreviate [, FirstDayOfWeek] ])=>String

Аналог:

```
function weekdayName(weekday, abbreviate = false, firstDayOfWeek = 1) {
    const fullNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    const shortNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    // Сдвигаем массив в зависимости от первого дня недели
    const shiftedFullNames = fullNames.slice(firstDayOfWeek - 1).concat(fullNames.slice(0, firstDayOfWeek - 1));
    const shiftedShortNames = 
          shortNames.slice(firstDayOfWeek - 1).concat(shortNames.slice(0, firstDayOfWeek - 1));
    if (weekday < 1 || weekday > 7) {
        throw new Error("Weekday must be between 1 and 7");
    }
    return abbreviate ? shiftedShortNames[weekday - 1] : shiftedFullNames[weekday - 1];
}
```

Weekday(Date,[FirstDayOfWeek])=>Integer

Аналог:

```
function weekday(date, firstDayOfWeek = 0) {
// {Date} date - Дата, для которой нужно определить номер дня недели.
// [firstDayOfWeek=0] - Первый день недели (0 для воскресенья)
    // Проверка на корректность первого дня недели
    if (firstDayOfWeek < 0 || firstDayOfWeek > 6) {
        throw new Error("firstDayOfWeek must be between 0 and 6");
    }
    
    // Получаем номер дня недели, где 0 - воскресенье, 1 - понедельник, и т.д.
    const day = date.getDay();
    // Сдвигаем номер дня недели в зависимости от первого дня недели
    const adjustedDay = (day - firstDayOfWeek + 7) % 7;
    // Возвращаем номер дня недели, где 1 - первый день недели, 2 - второй, и т.д.
    return adjustedDay + 1;
}
// Примеры использования
console.log(weekday(new Date('2024-07-10'))); // Например, 3 (если неделя начинается с воскресенья)
console.log(weekday(new Date('2024-07-10'), 1)); // Например, 2 (если неделя начинается с понедельника)
```

TimeValue(Time)=>Date

Аналог:

```
function timeValue(timeStr) {
    // Проверка на корректность входной строки
    if (typeof timeStr !== 'string') {
        throw new Error("Input must be a string");
    }
    // Регулярное выражение для проверки и разбора строки времени
    const timeRegex = /^(\d{2}):(\d{2})(:(\d{2}))?$/;
    const match = timeStr.match(timeRegex);
    if (!match) {
        throw new Error("Invalid time format. Use 'HH:MM:SS' or 'HH:MM'");
    }
    const hours = parseInt(match[1], 10);
    const minutes = parseInt(match[2], 10);
    const seconds = match[4] ? parseInt(match[4], 10) : 0;
    // Создаем объект Date для текущей даты, но с заданным временем
    const now = new Date();
    now.setHours(hours, minutes, seconds, 0);
    return now;
}
// Примеры использования
console.log(timeValue("14:30:00")); // Объект Date с временем 14:30:00 на текущей дате
console.log(timeValue("09:15"));    // Объект Date с временем 09:15:00 на текущей дате
```

TimeSerial(Hour,Minute,Second)=>Date

Аналог:

```
function timeSerial(hour, minute, second) {
    // Проверка валидности входных значений
    if (typeof hour !== 'number' || typeof minute !== 'number' || typeof second !== 'number') {
        throw new Error("Hour, minute, and second must be numbers");
    }
    if (hour < 0 || hour > 23) {
        throw new Error("Hour must be between 0 and 23");
    }
    if (minute < 0 || minute > 59) {
        throw new Error("Minute must be between 0 and 59");
    }
    if (second < 0 || second > 59) {
        throw new Error("Second must be between 0 and 59");
    }
    // Создаем объект Date для текущей даты и устанавливаем заданное время
    const now = new Date();
    now.setHours(hour, minute, second, 0);
    return now;
}
// Примеры использования
console.log(timeSerial(14, 30, 0)); // Объект Date с временем 14:30:00 на текущей дате
console.log(timeSerial(9, 15, 45)); // Объект Date с временем 09:15:45 на текущей дате
```

Timer()=>Single

Аналог:

```
function timer() {
    const now = new Date();
    const midnight = new Date(now.toDateString()); // Создает объект Date на полночь текущего дня
    const secondsSinceMidnight = (now - midnight) / 1000; // Разница в миллисекундах преобразована в секунды
    return secondsSinceMidnight;
}
// Пример использования
console.log(timer()); // Количество секунд, прошедших с полуночи
```

Time()=>Date

Аналог:

```
function currentTime() {
//String
    const now = new Date();
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');
    return `${hours}:${minutes}:${seconds}`;
}
```

Second(Time)=>Integer

Аналог:

```
ffunction second(time) {    return time.getSeconds();}
//Пример
const customTime = new Date(2023, 0, 1, 12, 30, 45); // 1 января 2023 года, 12:30:45
console.log(second(customTime)); // Выводит 45 (секунды)
```

Now()=>Date

Аналог:

```
function currentDateTime() {
// Текущая дата и время в формате "ГГГГ-ММ-ДД ЧЧ:ММ:СС"
    const now = new Date();
    // Получение частей даты и времени
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0'); // Месяцы начинаются с 0
    const day = String(now.getDate()).padStart(2, '0');
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');
    // Форматирование в строку
    return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}
```

MonthName(Month[, Abbreviate])=>String

Аналог:

```
function monthName(month, abbreviate = false) {
    // Проверка на допустимость номера месяца
    if (month < 1 || month > 12) {
        throw new Error('Номер месяца должен быть в диапазоне от 1 до 12.');
    }
    // Массив полных названий месяцев
    const fullMonthNames = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ];
    // Массив сокращённых названий месяцев
    const abbreviatedMonthNames = [
        'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
        'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
    ];
    // Выбор массива названий в зависимости от параметра abbreviate
    const monthNames = abbreviate ? abbreviatedMonthNames : fullMonthNames;
    return monthNames[month - 1];
}
```

Month(Date)=>Integer

Аналог:

```
function getMonth(date) {
    if (!(date instanceof Date)) {
        throw new TypeError('Аргумент должен быть объектом Date.');
    }
    // Месяцы в JavaScript индексируются с 0, поэтому добавляем 1
    return date.getMonth() + 1;  //   1..12
}
```

Minute(Time)=>Integer

Аналог:

```
function getMinutes(date) {
    if (!(date instanceof Date)) {
        throw new TypeError('Аргумент должен быть объектом Date.');
    }
    return date.getMinutes();
}
```

IsDate(Expression)=>Boolean

Аналог:

```
function isDate(value) {
    const date = new Date(value);
    // Проверяем, является ли дата допустимой и соответствует ли она исходному значению
    return !isNaN(date.getTime()) && value !== '';
}
```

Hour(Time)=>Integer

Аналог:

```
function getHour(time) {
    // Создаем объект Date, если time является строкой
    const date = (typeof time === 'string') ? new Date(`1970-01-01T${time}Z`) : new Date(time);
    // Проверяем, является ли date допустимым объектом Date
    if (isNaN(date.getTime())) {
        throw new Error('Invalid date or time value');
    }
    return date.getUTCHours(); // Используем getUTCHours для получения часа в пределах от 0 до 23
}
```

Day(Date)=>Integer

Аналог:

```
function getDay(date) {    return date.getDate();}
```

DateValue(Date)=>Date

Аналог (возможны проблемы из-за внутреннего формата дат):

```
function dateValue(dateString) {
    const dateObj = new Date(dateString);
    // Проверяем, является ли dateObj допустимым объектом Date
    if (isNaN(dateObj.getTime())) {
        throw new Error('Invalid date string');
    }
    // Устанавливаем время в 00:00:00, если нужно только дату без времени
    dateObj.setHours(0, 0, 0, 0);
    return dateObj;
}
// Примеры использования
    console.log(dateValue('2024-07-11')); // Thu Jul 11 2024 00:00:00 GMT+0000
    console.log(dateValue('07/11/2024')); // Thu Jul 11 2024 00:00:00 GMT+0000
    console.log(dateValue('11 Jul 2024')); // Thu Jul 11 2024 00:00:00 GMT+0000
```

DateSerial(Year,Month,Day)=>Date

Аналог:

```
function dateSerial(year, month, day) {
    // Поскольку в JavaScript месяцы начинаются с 0, мы вычитаем 1 из месяца
    const dateObj = new Date(year, month - 1, day);
    // Проверяем, является ли dateObj допустимым объектом Date
    if (isNaN(dateObj.getTime())) {
        throw new Error('Invalid date values');
    }
    return dateObj;
}
// Примеры использования
    console.log(dateSerial(2024, 7, 11)); // Thu Jul 11 2024 00:00:00 GMT+0000
    console.log(dateSerial(2024, 12, 31)); // Tue Dec 31 2024 00:00:00 GMT+0000
```

DatePart(Interval,Date,[FirstDayOfWeek],[FirstWeekOfYear])=>Integer

Аналог:

```
function datePart(interval, date, firstDayOfWeek = 0, firstWeekOfYear = 0) {
    switch (interval.toLowerCase()) {
        case 'yyyy':
            return date.getFullYear();
        case 'q':            // Определяем квартал по месяцу
            const month = date.getMonth();
            return Math.floor(month / 3) + 1;
        case 'm':
            return date.getMonth() + 1; // Месяцы в JavaScript начинаются с 0
        case 'd':
            return date.getDate();
        case 'w':
            return date.getDay(); // В JavaScript неделя начинается с воскресенья (0)
        case 'y':            // День года, т.е. сколько дней прошло с начала года
            const startOfYear = new Date(date.getFullYear(), 0, 1);
            return Math.ceil((date - startOfYear) / (24 * 60 * 60 * 1000)) + 1;
        case 'ww':            // Определение номера недели года
            const startOfYearWeek = new Date(date.getFullYear(), 0, 1);
            const weekDiff = (date - startOfYearWeek) / (7 * 24 * 60 * 60 * 1000);
            return Math.ceil(weekDiff) + 1;
        default:
            throw new Error('Invalid interval specified');
    }
}
// Примеры использования
const date = new Date(2024, 6, 11); // 11 июля 2024 года
console.log(datePart('yyyy', date)); // 2024
console.log(datePart('q', date)); // 3 (третий квартал)
console.log(datePart('m', date)); // 7 (июль)
console.log(datePart('d', date)); // 11 (день месяца)
console.log(datePart('w', date)); // 4 (четверг, неделя начинается с воскресенья)
console.log(datePart('y', date)); // 193 (день года)
console.log(datePart('ww', date)); // Номер недели года
```

DateDiff(Interval,Date1,Date2,[FirstDayOfWeek],[FirstWeekOfYear])=>  
Long

Аналог:

```
function dateDiff(interval, date1, date2, firstDayOfWeek = 0, firstWeekOfYear = 0) {
    // Убедимся, что date1 и date2 являются объектами Date
    if (!(date1 instanceof Date) || !(date2 instanceof Date)) {
        throw new Error('date1 and date2 must be Date objects');
    }
    // Вычисление разницы в миллисекундах
    const diffInMs = date2 - date1;    
    // Определяем интервал
    switch (interval.toLowerCase()) {
        case 'y': // Годы
            return date2.getFullYear() - date1.getFullYear();
        case 'm': // Месяцы
            return (date2.getFullYear() - date1.getFullYear()) * 12 + (date2.getMonth() - date1.getMonth());
        case 'd': // Дни
            return Math.floor(diffInMs / (1000 * 60 * 60 * 24));
        case 'h': // Часы
            return Math.floor(diffInMs / (1000 * 60 * 60));
        case 'n': // Минуты
            return Math.floor(diffInMs / (1000 * 60));
        case 's': // Секунды
            return Math.floor(diffInMs / 1000);
        default:
            throw new Error('Invalid interval specified');
    }
}
// Примеры использования
const date1 = new Date(2024, 0, 1); // 1 января 2024 года
const date2 = new Date(2024, 6, 11); // 11 июля 2024 года
console.log(dateDiff('y', date1, date2)); // 0 (разница в годах)
console.log(dateDiff('m', date1, date2)); // 6 (разница в месяцах)
console.log(dateDiff('d', date1, date2)); // 191 (разница в днях)
console.log(dateDiff('h', date1, date2)); // 4584 (разница в часах)
console.log(dateDiff('n', date1, date2)); // 275040 (разница в минутах)
console.log(dateDiff('s', date1, date2)); // 16502400 (разница в секундах)
```

DateAdd(Interval, Number, Date)=>Date

Аналог:

```
function dateAdd(interval, number, date) {
    // Убедимся, что date является объектом Date
    if (!(date instanceof Date)) {
        throw new Error('date must be a Date object');
    }
    // Создаем копию исходной даты
    const newDate = new Date(date);
    // Добавляем интервал в зависимости от типа
    switch (interval.toLowerCase()) {
        case 'days':
            newDate.setDate(newDate.getDate() + number);
            break;
        case 'months':
            newDate.setMonth(newDate.getMonth() + number);
            break;
        case 'years':
            newDate.setFullYear(newDate.getFullYear() + number);
            break;
        case 'hours':
            newDate.setHours(newDate.getHours() + number);
            break;
        case 'minutes':
            newDate.setMinutes(newDate.getMinutes() + number);
            break;
        case 'seconds':
            newDate.setSeconds(newDate.getSeconds() + number);
            break;
        default:
            throw new Error('Invalid interval specified');
    }
    return newDate;
}
// Примеры использования
const date = new Date(2024, 0, 1); // 1 января 2024 года
console.log(dateAdd('days', 10, date));    // 11 января 2024 года
console.log(dateAdd('months', 6, date));   // 1 июля 2024 года
console.log(dateAdd('years', 1, date));    // 1 января 2025 года
console.log(dateAdd('hours', 24, date));   // 2 января 2024 года
console.log(dateAdd('minutes', 120, date)); // 1 января 2024 года, 2 часа спустя
console.log(dateAdd('seconds', 3600, date)); // 1 января 2024 года, 1 час спустя
```

Date()=>Date

Аналог:

```
function getCurrentDate() {
    const now = new Date();
    // Устанавливаем время на полночь
    now.setHours(0, 0, 0, 0);
    return now;
}
```

Функции работы с указателями (3)

VarPtr(Ptr)=>Long

Аналог (функциональный аналог, так как нет доступа непосредственно к адресу объекта):

```
function getObjectReference(obj) {    // Возвращает ссылку на объект
    return obj;
}
```

ИЛИ через WeakMap

```
const weakMap = new WeakMap();
function storeObject(obj) {
    // Использование WeakMap для хранения объекта
    weakMap.set(obj, true);
}
// Пример использования
const myObject = { value: 42 };
storeObject(myObject);
// Проверка, существует ли объект в WeakMap
console.log(weakMap.has(myObject)); // true
```

StrPtr(Ptr)=>Long

Аналога нет (Строки не изменяемы, при изменении создается новая).

```
const str = "Hello, World!";
const anotherStr = str;
console.log(anotherStr); // "Hello, World!"
```

Для контроля строк можно использовать:  
1.Map

```
const stringMap = new Map();
function storeString(key, value) {
    stringMap.set(key, value);
}
// Пример использования
const key = "uniqueKey";
const value = "Hello, World!";
storeString(key, value);
// Получение строки из Map
console.log(stringMap.get(key)); // "Hello, World!"
```

2.Идентификаторы

```
const strings = {
    first: "Hello",
    second: "World"
};
const identifier = "first";
console.log(strings[identifier]); // "Hello"
```

3.Строковый объект

```
const strObj = new String("Hello, World!");
console.log(strObj.toString()); // "Hello, World!"
```

ObjPtr(Ptr)=>Long

Аналога нет:  
Для контроля объектов можно использовать:  
1.Map

```
const objectMap = new Map();
function storeObject(key, obj) {
    objectMap.set(key, obj);
}
// Пример использования
const key = "uniqueObject";
const obj = { name: "Example" };
storeObject(key, obj);
// Получение объекта из Map
console.log(objectMap.get(key)); // { name: "Example" }
```

2.Ссылки

```
const obj1 = { name: "Object1" };
const obj2 = obj1; // obj2 и obj1 указывают на один и тот же объект
console.log(obj2.name); // "Object1"
```

3. WeakMap

```
const weakMap = new WeakMap();
function storeObject(key, obj) {
    weakMap.set(key, obj);
}
// Пример использования
const key = {};
const obj = { name: "WeakObject" };
storeObject(key, obj);
// Получение объекта из WeakMap
console.log(weakMap.get(key)); // { name: "WeakObject" }

```

4.Метаданные

```
const objects = {};
function addMetadata(id, obj) {
    objects[id] = obj;
}
// Пример использования
const id = "metaObject";
const obj = { name: "MetadataObject" };
addMetadata(id, obj);
// Получение объекта из метаданных
console.log(objects[id]); // { name: "MetadataObject" }
```

Функции загрузки и отображения данных (8)

MsgBox(Prompt, [Buttons], [Title], [HelpFile], [Context])

Аналоги (приблизительные):

```
// alert - аналог MsgBox с одной кнопкой "OK"
alert("Это сообщение для пользователя.");

// confirm - аналог MsgBox с кнопками "OK" и "Cancel"
const userResponse = confirm("Вы уверены, что хотите продолжить?");
if (userResponse) {
    console.log("Пользователь нажал OK.");
} else {
    console.log("Пользователь нажал Cancel.");
}

// prompt - аналог MsgBox с полем ввода
const userInput = prompt("Введите ваше имя:", "Имя пользователя");
if (userInput !== null) {
    console.log(`Пользователь ввел: ${userInput}`);
} else {
    console.log("Пользователь отменил ввод.");
}
```

LoadResString(id)=>String

Аналога нет. Можно воспользоваться файлами json.  
Пример:  
Файл resources.json

```
{
    "1001": "Hello, World!",
    "1002": "Welcome to our application.",
    "1003": "An error occurred. Please try again later."
}
```

Загрузка из файла:

```
fetch('resources.json')
    .then(response => response.json())
    .then(data => {
        // Функция для получения строки ресурса по идентификатору
        function loadResString(id) {
            return data[id] || "Resource not found.";
        }

        // Пример использования функции
        const message = loadResString("1001");
        console.log(message); // Выведет: Hello, World!
    })
    .catch(error => console.error('Error loading resources:', error));
```

LoadResPicture(id, restype)

Аналога нет. Можно воспользоваться файлами json.  
Пример:  
Файл resources.json

```
{
    "2001": "/images/image1.png",
    "2002": "/images/image2.jpg"
}
```

Загрузка из файла по идентификатору:

```
function loadResPicture(id) {
    return fetch('resources.json')
        .then(response => response.json())
        .then(data => {
            const imagePath = data[id];
            if (imagePath) {
                return fetch(imagePath)
                    .then(response => response.blob())
                    .then(blob => URL.createObjectURL(blob));
            }
            throw new Error('Resource not found.');
        })
        .catch(error => console.error('Error loading resource:', error));
}

// Пример использования функции
loadResPicture("2001")
    .then(imgUrl => {
        document.getElementById('myImage').src = imgUrl;
    });
```

LoadResData(id,type)

Аналога нет. Можно воспользоваться файлами json.  
Пример:  
Файл resources.json

```
{
    "data1": {
        "type": "text",
        "path": "/resources/data1.txt"
    },
    "data2": {
        "type": "json",
        "path": "/resources/data2.json"
    }
}
```

Загрузка из файла:

```
Загрузка из файла
async function loadResData(id, type) {
    try {
        // Загрузка метаданных о ресурсах
        const response = await fetch('/path/to/resources.json');
        const resources = await response.json();
        const resource = resources[id];
        if (resource && resource.type === type) {
            const dataResponse = await fetch(resource.path);
            if (type === 'text') {
                return await dataResponse.text();
            } else if (type === 'json') {
                return await dataResponse.json();
            }
        }
        throw new Error('Resource not found or type mismatch.');
    } catch (error) {
        console.error('Error loading resource:', error);
    }
}

// Пример использования функции
loadResData('data1', 'text').then(data => {
    console.log(data);  // Вывод текста из ресурса
});
loadResData('data2', 'json').then(data => {
    console.log(data);  // Вывод данных из JSON
});
```

LoadPicture([FileName], [Size], [ColorDepth],[X,Y])

Аналога нет.

InputBox(Prompt[,Title] [,Default] [,XPos] [,YPos] [,HelpFile,Context])=>  
String

Аналог (приблизительный):

```
let userInput = prompt("Введите текст:", "Значение по умолчанию");
console.log("Введенный текст:", userInput);
```

IIf(Expression, TruePart, FalsePart)=>Part

Аналоги:

```
// 1.Тернарный оператор
let result = (number % 2 === 0) ? "Even" : "Odd";
// 2.Функция
function iIf(condition, truePart, falsePart) {    return condition ? truePart : falsePart;}
```

Choose(Index,item1 [, item2 [ ,..., [ itemN]] )=>item

Аналог:

```
let result = choose(index, 'Apple', 'Banana', 'Cherry');
```




---
