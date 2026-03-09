(function () {
    "use strict";

    // --- Conversion data ---
    var units = {
        length: {
            label: "Length",
            base: "m",
            units: {
                "mm": { name: "Millimeters", factor: 0.001 },
                "cm": { name: "Centimeters", factor: 0.01 },
                "m": { name: "Meters", factor: 1 },
                "km": { name: "Kilometers", factor: 1000 },
                "in": { name: "Inches", factor: 0.0254 },
                "ft": { name: "Feet", factor: 0.3048 },
                "yd": { name: "Yards", factor: 0.9144 },
                "mi": { name: "Miles", factor: 1609.344 }
            }
        },
        weight: {
            label: "Weight",
            base: "kg",
            units: {
                "mg": { name: "Milligrams", factor: 0.000001 },
                "g": { name: "Grams", factor: 0.001 },
                "kg": { name: "Kilograms", factor: 1 },
                "t": { name: "Tons", factor: 1000 },
                "oz": { name: "Ounces", factor: 0.0283495 },
                "lb": { name: "Pounds", factor: 0.453592 }
            }
        },
        temperature: {
            label: "Temperature",
            base: "C",
            units: {
                "C": { name: "Celsius" },
                "F": { name: "Fahrenheit" },
                "K": { name: "Kelvin" }
            },
            custom: true
        },
        area: {
            label: "Area",
            base: "m2",
            units: {
                "mm2": { name: "mm²", factor: 0.000001 },
                "cm2": { name: "cm²", factor: 0.0001 },
                "m2": { name: "m²", factor: 1 },
                "km2": { name: "km²", factor: 1000000 },
                "ha": { name: "Hectares", factor: 10000 },
                "ac": { name: "Acres", factor: 4046.8564 },
                "ft2": { name: "ft²", factor: 0.092903 }
            }
        },
        speed: {
            label: "Speed",
            base: "ms",
            units: {
                "ms": { name: "m/s", factor: 1 },
                "kmh": { name: "km/h", factor: 0.277778 },
                "mph": { name: "mph", factor: 0.44704 },
                "kn": { name: "Knots", factor: 0.514444 }
            }
        },
        data: {
            label: "Data",
            base: "B",
            units: {
                "B": { name: "Bytes", factor: 1 },
                "KB": { name: "Kilobytes", factor: 1024 },
                "MB": { name: "Megabytes", factor: 1048576 },
                "GB": { name: "Gigabytes", factor: 1073741824 },
                "TB": { name: "Terabytes", factor: 1099511627776 }
            }
        }
    };

    var currentValue = null;

    function convertTemperature(val, from, to) {
        if (from === to) return val;
        // Convert to Celsius first
        var celsius;
        switch (from) {
            case "C": celsius = val; break;
            case "F": celsius = (val - 32) * 5 / 9; break;
            case "K": celsius = val - 273.15; break;
            default: celsius = val;
        }
        // Convert from Celsius to target
        switch (to) {
            case "C": return celsius;
            case "F": return celsius * 9 / 5 + 32;
            case "K": return celsius + 273.15;
            default: return celsius;
        }
    }

    function convert(val, category, from, to) {
        if (from === to) return val;
        var cat = units[category];
        if (!cat) return NaN;

        if (cat.custom) {
            return convertTemperature(val, from, to);
        }

        var fromFactor = cat.units[from] ? cat.units[from].factor : 1;
        var toFactor = cat.units[to] ? cat.units[to].factor : 1;
        return val * fromFactor / toFactor;
    }

    function formatResult(num) {
        if (isNaN(num)) return "—";
        if (Math.abs(num) >= 1e9 || (Math.abs(num) < 0.0001 && num !== 0)) {
            return num.toExponential(4);
        }
        return parseFloat(num.toPrecision(10)).toString();
    }

    function populateUnits(category) {
        var cat = units[category];
        if (!cat) return;

        var fromSel = document.getElementById("unitFrom");
        var toSel = document.getElementById("unitTo");
        fromSel.innerHTML = "";
        toSel.innerHTML = "";

        var keys = Object.keys(cat.units);
        for (var i = 0; i < keys.length; i++) {
            var key = keys[i];
            var unit = cat.units[key];
            var opt1 = document.createElement("option");
            opt1.value = key;
            opt1.textContent = unit.name + " (" + key + ")";
            fromSel.appendChild(opt1);

            var opt2 = document.createElement("option");
            opt2.value = key;
            opt2.textContent = unit.name + " (" + key + ")";
            toSel.appendChild(opt2);
        }

        if (keys.length > 1) {
            toSel.selectedIndex = 1;
        }

        updateResult();
    }

    function updateResult() {
        var resultEl = document.getElementById("resultValue");
        if (currentValue === null || isNaN(currentValue)) {
            resultEl.textContent = "—";
            return;
        }

        var category = document.getElementById("category").value;
        var from = document.getElementById("unitFrom").value;
        var to = document.getElementById("unitTo").value;
        var result = convert(currentValue, category, from, to);
        resultEl.textContent = formatResult(result);
    }

    function showStatus(msg, isError) {
        var el = document.getElementById("statusMsg");
        el.textContent = msg;
        el.className = "status-msg" + (isError ? " error" : " success");
        setTimeout(function () {
            el.className = "status-msg hidden";
        }, 2000);
    }

    function readCurrentCell() {
        window.Asc.plugin.callCommand(function () {
            var oSheet = Api.GetActiveSheet();
            var oRange = oSheet.GetSelection();
            var val = "";
            oRange.ForEach(function (cell) {
                val = cell.GetValue();
            });
            return val;
        }, false, false, function (result) {
            var num = parseFloat(result);
            currentValue = isNaN(num) ? null : num;
            document.getElementById("currentValue").textContent =
                currentValue !== null ? currentValue.toString() : "—";
            updateResult();
        });
    }

    function writeResult(mode) {
        var category = document.getElementById("category").value;
        var from = document.getElementById("unitFrom").value;
        var to = document.getElementById("unitTo").value;
        if (currentValue === null) {
            showStatus("No numeric value selected", true);
            return;
        }
        var result = convert(currentValue, category, from, to);
        var resultStr = formatResult(result);
        var modeStr = mode;

        var fnBody = '(function() {\n' +
            'var oSheet = Api.GetActiveSheet();\n' +
            'var oRange = oSheet.GetSelection();\n' +
            'var mode = "' + modeStr + '";\n' +
            'var result = "' + resultStr + '";\n' +
            'if (mode === "replace") {\n' +
            '  oRange.ForEach(function(cell) {\n' +
            '    cell.SetValue(result);\n' +
            '  });\n' +
            '} else {\n' +
            '  var addr = oRange.GetAddress(true, true, "xlA1", false);\n' +
            '  var match = addr.match(/\\$([A-Z]+)\\$([0-9]+)/);\n' +
            '  if (match) {\n' +
            '    var colStr = match[1];\n' +
            '    var row = parseInt(match[2]);\n' +
            '    var colNum = 0;\n' +
            '    for (var c = 0; c < colStr.length; c++) {\n' +
            '      colNum = colNum * 26 + colStr.charCodeAt(c) - 64;\n' +
            '    }\n' +
            '    colNum += 1;\n' +
            '    var col = "";\n' +
            '    while (colNum > 0) {\n' +
            '      var rem = (colNum - 1) % 26;\n' +
            '      col = String.fromCharCode(65 + rem) + col;\n' +
            '      colNum = Math.floor((colNum - 1) / 26);\n' +
            '    }\n' +
            '    oSheet.GetRange(col + row).SetValue(result);\n' +
            '  }\n' +
            '}\n' +
            '})()';

        window.Asc.plugin.info.recalculate = true;
        window.Asc.plugin.callCommand(new Function(fnBody), false, false, function () {
            showStatus("✓ Done!", false);
        });
    }

    window.Asc.plugin.init = function () {
        var catSel = document.getElementById("category");
        var fromSel = document.getElementById("unitFrom");
        var toSel = document.getElementById("unitTo");

        catSel.addEventListener("change", function () {
            populateUnits(catSel.value);
        });
        fromSel.addEventListener("change", updateResult);
        toSel.addEventListener("change", updateResult);

        document.getElementById("swapBtn").addEventListener("click", function () {
            var tmp = fromSel.value;
            fromSel.value = toSel.value;
            toSel.value = tmp;
            updateResult();
        });

        document.getElementById("btnReplace").addEventListener("click", function () {
            writeResult("replace");
        });
        document.getElementById("btnInsertRight").addEventListener("click", function () {
            writeResult("right");
        });

        populateUnits("length");
        readCurrentCell();
    };

    window.Asc.plugin.button = function (id) {
        this.executeCommand("close", "");
    };

    // Re-read cell on selection change (initOnSelectionChanged is true)
    window.Asc.plugin.onExternalMouseUp = function () {
        readCurrentCell();
    };

    // Backup: also handle init being called again on selection change
    var origInit = window.Asc.plugin.init;
    window.Asc.plugin.init = function (data) {
        if (origInit._initialized) {
            readCurrentCell();
            return;
        }
        origInit._initialized = true;
        origInit.call(this, data);
    };

    Asc.plugin.onThemeChanged = function (theme) {
        window.Asc.plugin.onThemeChangedBase(theme);
        var rules = "";
        if (theme["background-toolbar"]) {
            rules += "body { background-color: " + theme["background-toolbar"] + "; }\n";
        }
        if (theme["text-normal"]) {
            rules += "body, select, .value-number, .result-value, .action-btn { color: " + theme["text-normal"] + "; }\n";
        }
        if (theme["background-normal"]) {
            rules += "select, .current-value-card, .result-card, .secondary-btn { background-color: " + theme["background-normal"] + "; }\n";
        }
        if (theme["border-regular-control"]) {
            rules += "select, .current-value-card, .result-card, .secondary-btn { border-color: " + theme["border-regular-control"] + "; }\n";
        }
        if (theme["text-tertiary"]) {
            rules += "label, .value-label, .result-label { color: " + theme["text-tertiary"] + "; }\n";
        }
        if (theme["background-primary-dialog-button"]) {
            rules += ".primary-btn { background-color: " + theme["background-primary-dialog-button"] + "; border-color: " + theme["background-primary-dialog-button"] + "; }\n";
        }
        if (theme["text-inverse"]) {
            rules += ".primary-btn { color: " + theme["text-inverse"] + "; }\n";
        }
        if (theme["highlight-primary-dialog-button-hover"]) {
            rules += ".primary-btn:hover { background-color: " + theme["highlight-primary-dialog-button-hover"] + "; }\n";
        }
        if (theme["highlight-button-hover"]) {
            rules += ".secondary-btn:hover, .swap-btn:hover { background-color: " + theme["highlight-button-hover"] + "; }\n";
        }
        var styleEl = document.getElementById("pluginStyles");
        if (styleEl) styleEl.innerHTML = rules;
    };
})();
