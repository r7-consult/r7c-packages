(function () {
    "use strict";

    // --- Data generation functions ---
    var firstNames = ["Alice", "Bob", "Carol", "David", "Eva", "Frank", "Grace", "Henry", "Iris", "Jack",
        "Kate", "Leo", "Mia", "Noah", "Olivia", "Peter", "Quinn", "Rose", "Sam", "Tina",
        "Uma", "Victor", "Wendy", "Xander", "Yara", "Zach", "Anna", "Ben", "Clara", "Dan"];
    var lastNames = ["Smith", "Johnson", "Brown", "Davis", "Wilson", "Moore", "Taylor", "Anderson",
        "Thomas", "Jackson", "White", "Harris", "Martin", "Garcia", "Clark", "Lewis",
        "Robinson", "Walker", "Young", "Allen", "King", "Wright", "Scott", "Green", "Baker"];
    var domains = ["gmail.com", "yahoo.com", "outlook.com", "mail.com", "example.org", "company.net", "work.io"];
    var loremWords = ["lorem", "ipsum", "dolor", "sit", "amet", "consectetur", "adipiscing", "elit",
        "sed", "do", "eiusmod", "tempor", "incididunt", "ut", "labore", "et", "dolore",
        "magna", "aliqua", "enim", "ad", "minim", "veniam", "quis", "nostrud"];

    function randInt(min, max) {
        return Math.floor(Math.random() * (max - min + 1)) + min;
    }

    function randFloat(min, max) {
        return (Math.random() * (max - min) + min).toFixed(2);
    }

    function uuid4() {
        return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, function (c) {
            var r = Math.random() * 16 | 0;
            var v = c === "x" ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }

    function randDate(fromStr, toStr) {
        var from = new Date(fromStr).getTime();
        var to = new Date(toStr).getTime();
        var d = new Date(from + Math.random() * (to - from));
        return d.toISOString().split("T")[0];
    }

    function randName() {
        return firstNames[randInt(0, firstNames.length - 1)] + " " + lastNames[randInt(0, lastNames.length - 1)];
    }

    function randEmail() {
        var fn = firstNames[randInt(0, firstNames.length - 1)].toLowerCase();
        var ln = lastNames[randInt(0, lastNames.length - 1)].toLowerCase();
        var dom = domains[randInt(0, domains.length - 1)];
        return fn + "." + ln + randInt(1, 99) + "@" + dom;
    }

    function randPassword(len) {
        var chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%&*";
        var result = "";
        for (var i = 0; i < (len || 12); i++) {
            result += chars.charAt(randInt(0, chars.length - 1));
        }
        return result;
    }

    function randLorem() {
        var count = randInt(5, 15);
        var words = [];
        for (var i = 0; i < count; i++) {
            words.push(loremWords[randInt(0, loremWords.length - 1)]);
        }
        var sentence = words.join(" ");
        return sentence.charAt(0).toUpperCase() + sentence.slice(1) + ".";
    }

    function generateData(type, count, params) {
        var data = [];
        for (var i = 0; i < count; i++) {
            var val;
            switch (type) {
                case "integers":
                    val = String(randInt(params.min, params.max));
                    break;
                case "decimals":
                    val = randFloat(params.min, params.max);
                    break;
                case "sequence":
                    val = String(params.start + i * params.step);
                    break;
                case "dates":
                    val = randDate(params.dateFrom, params.dateTo);
                    break;
                case "uuid":
                    val = uuid4();
                    break;
                case "names":
                    val = randName();
                    break;
                case "emails":
                    val = randEmail();
                    break;
                case "passwords":
                    val = randPassword(14);
                    break;
                case "lorem":
                    val = randLorem();
                    break;
                default:
                    val = String(randInt(1, 100));
            }
            data.push(val);
        }
        return data;
    }

    function showStatus(msg, isError) {
        var el = document.getElementById("statusMsg");
        el.textContent = msg;
        el.className = "status-msg" + (isError ? " error" : " success");
        setTimeout(function () {
            el.className = "status-msg hidden";
        }, 2500);
    }

    // --- UI logic ---
    function updateFormVisibility() {
        var type = document.getElementById("dataType").value;
        document.getElementById("rangeGroup").style.display =
            (type === "integers" || type === "decimals") ? "" : "none";
        document.getElementById("seqGroup").style.display =
            type === "sequence" ? "" : "none";
        document.getElementById("dateRangeGroup").style.display =
            type === "dates" ? "" : "none";
    }

    function insertData() {
        var type = document.getElementById("dataType").value;
        var count = parseInt(document.getElementById("rowCount").value) || 10;
        if (count < 1) count = 1;
        if (count > 10000) count = 10000;

        var direction = document.querySelector('input[name="direction"]:checked').value;
        var params = {
            min: parseInt(document.getElementById("rangeMin").value) || 0,
            max: parseInt(document.getElementById("rangeMax").value) || 100,
            start: parseInt(document.getElementById("seqStart").value) || 1,
            step: parseInt(document.getElementById("seqStep").value) || 1,
            dateFrom: document.getElementById("dateFrom").value || "2024-01-01",
            dateTo: document.getElementById("dateTo").value || "2025-12-31"
        };

        var data = generateData(type, count, params);
        var dataJson = JSON.stringify(data);
        var dirStr = direction;

        var fnBody = '(function() {\n' +
            'var oSheet = Api.GetActiveSheet();\n' +
            'var oCell = oSheet.GetSelection();\n' +
            'var addr = oCell.GetAddress(true, true, "xlA1", false);\n' +
            'var match = addr.match(/\\$([A-Z]+)\\$([0-9]+)/);\n' +
            'if (!match) return;\n' +
            'var colStr = match[1];\n' +
            'var rowStart = parseInt(match[2]);\n' +
            'var data = ' + dataJson + ';\n' +
            'var dir = "' + dirStr + '";\n' +
            'for (var i = 0; i < data.length; i++) {\n' +
            '  var cellRef;\n' +
            '  if (dir === "down") {\n' +
            '    cellRef = colStr + (rowStart + i);\n' +
            '  } else {\n' +
            '    var colNum = 0;\n' +
            '    for (var c = 0; c < colStr.length; c++) {\n' +
            '      colNum = colNum * 26 + colStr.charCodeAt(c) - 64;\n' +
            '    }\n' +
            '    colNum += i;\n' +
            '    var col = "";\n' +
            '    while (colNum > 0) {\n' +
            '      var rem = (colNum - 1) % 26;\n' +
            '      col = String.fromCharCode(65 + rem) + col;\n' +
            '      colNum = Math.floor((colNum - 1) / 26);\n' +
            '    }\n' +
            '    cellRef = col + rowStart;\n' +
            '  }\n' +
            '  oSheet.GetRange(cellRef).SetValue(data[i]);\n' +
            '}\n' +
            '})()';

        window.Asc.plugin.info.recalculate = true;
        window.Asc.plugin.callCommand(new Function(fnBody), false, false, function () {
            showStatus("✓ Generated " + data.length + " values!", false);
        });
    }

    window.Asc.plugin.init = function () {
        document.getElementById("dataType").addEventListener("change", updateFormVisibility);
        document.getElementById("btnGenerate").addEventListener("click", insertData);
        updateFormVisibility();
    };

    window.Asc.plugin.button = function (id) {
        this.executeCommand("close", "");
    };

    Asc.plugin.onThemeChanged = function (theme) {
        window.Asc.plugin.onThemeChangedBase(theme);
        var rules = "";
        if (theme["background-toolbar"]) {
            rules += "body { background-color: " + theme["background-toolbar"] + "; }\n";
        }
        if (theme["text-normal"]) {
            rules += "body, select, input, .radio-label { color: " + theme["text-normal"] + "; }\n";
        }
        if (theme["background-normal"]) {
            rules += "select, input[type='number'], input[type='date'], .generate-btn { background-color: " + theme["background-normal"] + "; }\n";
        }
        if (theme["border-regular-control"]) {
            rules += "select, input[type='number'], input[type='date'], .generate-btn { border-color: " + theme["border-regular-control"] + "; }\n";
        }
        if (theme["text-tertiary"]) {
            rules += "label, .range-separator { color: " + theme["text-tertiary"] + "; }\n";
        }
        if (theme["background-primary-dialog-button"]) {
            rules += ".generate-btn { background-color: " + theme["background-primary-dialog-button"] + "; }\n";
        }
        if (theme["text-inverse"]) {
            rules += ".generate-btn { color: " + theme["text-inverse"] + "; }\n";
        }
        if (theme["highlight-primary-dialog-button-hover"]) {
            rules += ".generate-btn:hover { background-color: " + theme["highlight-primary-dialog-button-hover"] + "; }\n";
        }
        var styleEl = document.getElementById("pluginStyles");
        if (styleEl) styleEl.innerHTML = rules;
    };
})();
