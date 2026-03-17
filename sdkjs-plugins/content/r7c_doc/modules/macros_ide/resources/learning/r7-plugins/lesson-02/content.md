# Занятие 1: Основы JavaScript

JavaScript — это язык программирования, используемый для создания интерактивных веб-страниц. Он позволяет реализовывать сложные элементы веб-страниц, такие как динамические обновления контента, интерактивные карты и анимации.

 `var name = "John"; let age = 25; const isStudent = true;`

```
var name = "John";
let age = 25;
const isStudent = true;

```

 `let number = 42; let string = "Hello, World!"; let boolean = true; let array = [1, 2, 3]; let object = {name: "John", age: 25}; let undefinedVariable; let nullVariable = null;`

```
let number = 42;
let string = "Hello, World!";
let boolean = true;
let array = [1, 2, 3];
let object = {name: "John", age: 25};
let undefinedVariable;
let nullVariable = null;

```

 `let fruits = ["Apple", "Banana", "Cherry"]; fruits.push("Orange"); // Добавляет элемент в конец массива console.log(fruits); // ["Apple", "Banana", "Cherry", "Orange"] let firstFruit = fruits.shift(); // Удаляет первый элемент массива console.log(firstFruit); // "Apple" console.log(fruits); // ["Banana", "Cherry", "Orange"]`

```
let fruits = ["Apple", "Banana", "Cherry"];
fruits.push("Orange");  // Добавляет элемент в конец массива
console.log(fruits);  // ["Apple", "Banana", "Cherry", "Orange"]
let firstFruit = fruits.shift();  // Удаляет первый элемент массива
console.log(firstFruit);  // "Apple"
console.log(fruits);  // ["Banana", "Cherry", "Orange"]

```

 `//Определение и вызов функций function greet(name) { return "Hello, " + name + "!"; } console.log(greet("Alice")); // "Hello, Alice!" //Функции стрелки (Arrow Functions) const add = (a, b) => a + b; console.log(add(2, 3)); // 5`

```
//Определение и вызов функций
function greet(name) {
    return "Hello, " + name + "!";
}
console.log(greet("Alice"));  // "Hello, Alice!"

//Функции стрелки (Arrow Functions)
const add = (a, b) => a + b;
console.log(add(2, 3));  // 5

```

 `//Условные операторы if, else if, else let age = 18; if (age < 18) { console.log("You are a minor."); } else if (age < 21) { console.log("You are a young adult."); } else { console.log("You are an adult."); } //Тернарный оператор let isMember = true; let price = isMember ? 10 : 20; console.log(price); // 10`

```
//Условные операторы if, else if, else
let age = 18;
if (age < 18) {
    console.log("You are a minor.");
} else if (age < 21) {
    console.log("You are a young adult.");
} else {
    console.log("You are an adult.");
}

//Тернарный оператор
let isMember = true;
let price = isMember ? 10 : 20;
console.log(price);  // 10

```

 `//for for (let i = 0; i < 5; i++) { console.log(i); } //while let i = 0; while (i < 5) { console.log(i); i++; } //do...while let i = 0; do { console.log(i); i++; } while (i < 5);`

```
//for
for (let i = 0; i < 5; i++) {
    console.log(i);
}

//while
let i = 0;
while (i < 5) {
    console.log(i);
    i++;
}

//do...while
let i = 0;
do {
    console.log(i);
    i++;
} while (i < 5);

```

 `//Поиск элементов let element = document.getElementById("myElement"); let elements = document.getElementsByClassName("myClass"); let elements = document.getElementsByTagName("p"); let element = document.querySelector(".myClass"); let elements = document.querySelectorAll(".myClass"); //Изменение элементов let element = document.getElementById("myElement"); element.textContent = "New Text"; element.style.color = "red"; element.setAttribute("data-custom", "value"); //Создание и удаление элементов. let newElement = document.createElement('div'); newElement.innerHTML = 'Hello, World!'; document.body.appendChild(newElement); document.body.removeChild(newElement);`

```
//Поиск элементов
let element = document.getElementById("myElement");
let elements = document.getElementsByClassName("myClass");
let elements = document.getElementsByTagName("p");
let element = document.querySelector(".myClass");
let elements = document.querySelectorAll(".myClass");

//Изменение элементов
let element = document.getElementById("myElement");
element.textContent = "New Text";
element.style.color = "red";
element.setAttribute("data-custom", "value");

//Создание и удаление элементов.
let newElement = document.createElement('div');
newElement.innerHTML = 'Hello, World!';
document.body.appendChild(newElement);
document.body.removeChild(newElement);

```

JavaScript позволяет реагировать на действия пользователей через обработчики событий. Вот как вы можете добавить обработчик события к элементу:

Используйте методы addEventListener и removeEventListener для назначения и удаления обработчиков событий:

 `<button id="myButton">Click me</button> <script> document.getElementById('myButton').addEventListener('click', function() { alert('Button clicked!'); }); </script>`

```
<button id="myButton">Click me</button>
<script>
document.getElementById('myButton').addEventListener('click', function() {
    		alert('Button clicked!');
});
</script>

```

click — клик по элементу

mouseover — наведение курсора на элемент

mouseout — увод курсора с элемента

keydown — нажатие клавиши

load — полная загрузка страницы

-        click — клик по элементу
-        mouseover — наведение курсора на элемент
-        mouseout — увод курсора с элемента
-        keydown — нажатие клавиши
-        load — полная загрузка страницы

 `<form id="myForm"> <input type="text" id="myInput" /> <button type="submit">Submit</button> </form> <script> document.getElementById('myForm').addEventListener('submit', function(event) { event.preventDefault(); alert(document.getElementById('myInput').value); }); </script>`

Вы можете реагировать на события формы, такие как отправка

```
<form id="myForm">
    <input type="text" id="myInput" />
    <button type="submit">Submit</button>
</form>

<script>
document.getElementById('myForm').addEventListener('submit', function(event) {
    		event.preventDefault();
   		 alert(document.getElementById('myInput').value);
});
</script>

```

Асинхронное программирование позволяет выполнять задачи без блокировки основного потока. Основные механизмы включают коллбэки, промисы и async/await.

 `function fetchData(callback) { setTimeout(() => { callback('Data received'); }, 1000); } fetchData(function(data) { console.log(data); });`

Функции обратного вызова выполняются после завершения асинхронной операции.

```
function fetchData(callback) {
    setTimeout(() => {
        callback('Data received');
    }, 1000);
}

fetchData(function(data) {
    console.log(data);
});

```

 `let promise = new Promise((resolve, reject) => { let success = true; if (success) { resolve('Operation succeeded'); } else { reject('Operation failed'); } }); promise.then((message) => { console.log(message); }).catch((error) => { console.error(error); });`

Промисы обязательно содержат в себе колбэки. Но сами по себе это отдельный класс: Promise – это специальный объект, который содержит своё состояние. Вначале pending («ожидание»), затем – одно из: fulfilled («выполнено успешно») или rejected («выполнено с ошибкой»). На promise можно навешивать колбэки двух типов: onFulfilled – срабатывают, когда promise в состоянии «выполнен успешно». onRejected – срабатывают, когда promise в состоянии «выполнен с ошибкой». Код, которому надо сделать что-то асинхронно, создаёт объект promise и возвращает его. Внешний код, получив promise, навешивает на него обработчики. По завершении процесса асинхронный код переводит promise в состояние fulfilled (с результатом) или rejected (с ошибкой). При этом автоматически вызываются соответствующие обработчики во внешнем коде..

```
let promise = new Promise((resolve, reject) => {
    let success = true;
    if (success) {
        resolve('Operation succeeded');
    } else {
        reject('Operation failed');
    }
});

promise.then((message) => {
    console.log(message);
}).catch((error) => {
    console.error(error);
});

```

 `async function fetchData() { try { let response = await fetch('https://jsonplaceholder.typicode.com/posts/1'); let data = await response.json(); console.log(data); } catch (error) { console.error('Error:', error); } } fetchData();`

Async/await делает асинхронный код более читаемым и управляемым.

```
async function fetchData() {
    try {
        let response = await fetch('https://jsonplaceholder.typicode.com/posts/1');
        let data = await response.json();
        console.log(data);
    } catch (error) {
        console.error('Error:', error);
    }
}
fetchData();

```

 `//Экспорт и импорт модулей // file: math.js export function add(a, b) { return a + b; } // file: main.js import { add } from './math.js'; console.log(add(2, 3));`

Модули помогают организовать код, разделяя его на более мелкие, управляемые части

```
//Экспорт и импорт модулей
// file: math.js
export function add(a, b) {
    return a + b;
}

// file: main.js
import { add } from './math.js';
console.log(add(2, 3));

```

 `try { let result = riskyFunction(); } catch (error) { console.error('An error occurred:', error); } finally { console.log('This will always run'); }`

Используйте конструкции try...catch для обработки ошибок.

```
try {
    let result = riskyFunction();
} catch (error) {
    console.error('An error occurred:', error);
} finally {
    console.log('This will always run');
}

```

 `class CustomError extends Error { constructor(message) { super(message); this.name = 'CustomError'; } } throw new CustomError('Something went wrong');`

Вы можете создавать свои собственные ошибки.

```
class CustomError extends Error {
    constructor(message) {
        super(message);
        this.name = 'CustomError';
    }
}

throw new CustomError('Something went wrong');

```

 `function createCounter() { let count = 0; return function() { count++; return count; } } let counter = createCounter(); console.log(counter()); // 1 console.log(counter()); // 2`

Замыкания позволяют функции запоминать контекст, в котором они были созданы.

```
function createCounter() {
    let count = 0;
    return function() {
        count++;
        return count;
    }
}

let counter = createCounter();
console.log(counter()); // 1
console.log(counter()); // 2

```

 `//Классы и наследование class Animal { constructor(name) { this.name = name; } speak() { console.log(`${this.name} makes a noise.`); } } class Dog extends Animal { speak() { console.log(`${this.name} barks.`); } } let dog = new Dog('Rex'); dog.speak(); // Rex barks.`

JavaScript поддерживает объектно-ориентированное программирование, включая классы и наследование.

```
//Классы и наследование
class Animal {
    constructor(name) {
        this.name = name;
    }

    speak() {
        console.log(`${this.name} makes a noise.`);
    }
}

class Dog extends Animal {
    speak() {
        console.log(`${this.name} barks.`);
    }
}

let dog = new Dog('Rex');
dog.speak(); // Rex barks.

```

 `fetch('https://jsonplaceholder.typicode.com/posts') .then(response => response.json()) .then(data => console.log(data)) .catch(error => console.error('Error:', error));`

Fetch API используется для выполнения сетевых запросов.

```
fetch('https://jsonplaceholder.typicode.com/posts')
    .then(response => response.json())
    .then(data => console.log(data))
    .catch(error => console.error('Error:', error));

```

 `let xhr = new XMLHttpRequest(); xhr.open('GET', 'https://jsonplaceholder.typicode.com/posts', true); xhr.onload = function() { if (xhr.status >= 200 && xhr.status < 300) { console.log(JSON.parse(xhr.responseText)); } else { console.error('Request failed'); } }; xhr.send();`

Старый способ выполнения сетевых запросов.

```
let xhr = new XMLHttpRequest();
xhr.open('GET', 'https://jsonplaceholder.typicode.com/posts', true);
xhr.onload = function() {
    if (xhr.status >= 200 && xhr.status < 300) {
        console.log(JSON.parse(xhr.responseText));
    } else {
        console.error('Request failed');
    }
};
xhr.send();

```

 `function* generator() { yield 1; yield 2; yield 3; } let gen = generator(); console.log(gen.next().value); // 1 console.log(gen.next().value); // 2 console.log(gen.next().value); // 3`

Генераторы позволяют создавать функции, выполнение которых можно приостанавливать и возобновлять.

```
function* generator() {
    yield 1;
    yield 2;
    yield 3;
}

let gen = generator();
console.log(gen.next().value); // 1
console.log(gen.next().value); // 2
console.log(gen.next().value); // 3

```

 `let target = { message1: "hello", message2: "everyone" }; let handler = { get: function(target, prop, receiver) { if (prop === "message1") { return "world"; } return Reflect.get(...arguments); } };`

Proxy позволяет создавать объекты-посредники для других объектов, перехватывая операции с ними.

```
let target = {
    message1: "hello",
    message2: "everyone"
};

let handler = {
    get: function(target, prop, receiver) {
        if (prop === "message1") {
            return "world";
        }
        return Reflect.get(...arguments);
    }
};

```


---


