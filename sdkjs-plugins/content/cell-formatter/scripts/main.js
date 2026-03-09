(function () {
    "use strict";

    function showStatus(msg, isError) {
        var el = document.getElementById("statusMsg");
        el.textContent = msg;
        el.className = "status-msg" + (isError ? " error" : " success");
        setTimeout(function () {
            el.className = "status-msg hidden";
        }, 2000);
    }

    function transformCells(transformFn) {
        window.Asc.plugin.callCommand(function () {
            var oSheet = Api.GetActiveSheet();
            var oRange = oSheet.GetSelection();
            oRange.ForEach(function (cell) {
                var val = cell.GetValue();
                if (val !== "" && val !== null && val !== undefined) {
                    var newVal = transformFn(val);
                    if (newVal !== val) {
                        cell.SetValue(newVal);
                    }
                }
            });
        }, false, false, function () {
            showStatus("✓ Done!", false);
        });
    }

    // The transformFn must be serialized as a string and embedded in callCommand
    function applyTransform(type, extra) {
        window.Asc.plugin.callCommand(function () {
            var oSheet = Api.GetActiveSheet();
            var oRange = oSheet.GetSelection();
            var type = "%%TYPE%%";
            var extra = "%%EXTRA%%";

            oRange.ForEach(function (cell) {
                var val = cell.GetValue();
                if (val === "" || val === null || val === undefined) return;

                var newVal = val;
                switch (type) {
                    case "upper":
                        newVal = val.toUpperCase();
                        break;
                    case "lower":
                        newVal = val.toLowerCase();
                        break;
                    case "title":
                        newVal = val.replace(/\w\S*/g, function (txt) {
                            return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
                        });
                        break;
                    case "sentence":
                        newVal = val.charAt(0).toUpperCase() + val.slice(1).toLowerCase();
                        break;
                    case "trim":
                        newVal = val.replace(/\s+/g, " ").trim();
                        break;
                    case "clean":
                        newVal = val.replace(/[^\x20-\x7E\u00A0-\uFFFF]/g, "");
                        break;
                    case "reverse":
                        newVal = val.split("").reverse().join("");
                        break;
                    case "prefix":
                        newVal = extra + val;
                        break;
                    case "suffix":
                        newVal = val + extra;
                        break;
                }
                if (newVal !== val) {
                    cell.SetValue(newVal);
                }
            });
        }.toString()
            .replace("%%TYPE%%", type)
            .replace("%%EXTRA%%", (extra || "").replace(/"/g, '\\"')),
            false, false, function () {
                showStatus("✓ Done!", false);
            });
    }

    // Since callCommand takes a function, we need to build it properly
    function runTransform(type, extra) {
        var extraSafe = (extra || "").replace(/\\/g, "\\\\").replace(/"/g, '\\"');
        var fnBody = '(function() {\n' +
            'var oSheet = Api.GetActiveSheet();\n' +
            'var oRange = oSheet.GetSelection();\n' +
            'var type = "' + type + '";\n' +
            'var extra = "' + extraSafe + '";\n' +
            'oRange.ForEach(function(cell) {\n' +
            '  var val = cell.GetValue();\n' +
            '  if (val === "" || val === null || val === undefined) return;\n' +
            '  var newVal = val;\n' +
            '  switch(type) {\n' +
            '    case "upper": newVal = val.toUpperCase(); break;\n' +
            '    case "lower": newVal = val.toLowerCase(); break;\n' +
            '    case "title": newVal = val.replace(/\\w\\S*/g, function(t){ return t.charAt(0).toUpperCase()+t.substr(1).toLowerCase(); }); break;\n' +
            '    case "sentence": newVal = val.charAt(0).toUpperCase()+val.slice(1).toLowerCase(); break;\n' +
            '    case "trim": newVal = val.replace(/\\s+/g," ").trim(); break;\n' +
            '    case "clean": newVal = val.replace(/[^\\x20-\\x7E\\u00A0-\\uFFFF]/g,""); break;\n' +
            '    case "reverse": newVal = val.split("").reverse().join(""); break;\n' +
            '    case "prefix": newVal = extra + val; break;\n' +
            '    case "suffix": newVal = val + extra; break;\n' +
            '  }\n' +
            '  if (newVal !== val) cell.SetValue(newVal);\n' +
            '});\n' +
            '})()';

        window.Asc.plugin.info.recalculate = true;
        window.Asc.plugin.callCommand(new Function(fnBody), false, false, function () {
            showStatus("✓ Applied!", false);
        });
    }

    function runDedupe() {
        window.Asc.plugin.callCommand(function () {
            var oSheet = Api.GetActiveSheet();
            var oRange = oSheet.GetSelection();
            var values = [];
            var cells = [];

            oRange.ForEach(function (cell) {
                values.push(cell.GetValue());
                cells.push(cell);
            });

            var seen = {};
            for (var i = 0; i < values.length; i++) {
                if (seen[values[i]] && values[i] !== "") {
                    cells[i].SetValue("");
                } else {
                    seen[values[i]] = true;
                }
            }
        }, false, false, function () {
            showStatus("✓ Duplicates removed!", false);
        });
    }

    window.Asc.plugin.init = function () {
        // Button click handlers
        document.getElementById("btnUpper").onclick = function () { runTransform("upper"); };
        document.getElementById("btnLower").onclick = function () { runTransform("lower"); };
        document.getElementById("btnTitle").onclick = function () { runTransform("title"); };
        document.getElementById("btnSentence").onclick = function () { runTransform("sentence"); };
        document.getElementById("btnTrim").onclick = function () { runTransform("trim"); };
        document.getElementById("btnClean").onclick = function () { runTransform("clean"); };
        document.getElementById("btnReverse").onclick = function () { runTransform("reverse"); };
        document.getElementById("btnDedupe").onclick = function () { runDedupe(); };

        document.getElementById("btnAddPrefix").onclick = function () {
            var prefix = document.getElementById("inputPrefix").value;
            if (!prefix) { showStatus("Enter a prefix first", true); return; }
            runTransform("prefix", prefix);
        };
        document.getElementById("btnAddSuffix").onclick = function () {
            var suffix = document.getElementById("inputSuffix").value;
            if (!suffix) { showStatus("Enter a suffix first", true); return; }
            runTransform("suffix", suffix);
        };
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
            rules += "body, .action-btn, .text-input { color: " + theme["text-normal"] + "; }\n";
        }
        if (theme["background-normal"]) {
            rules += ".action-btn, .text-input { background-color: " + theme["background-normal"] + "; }\n";
        }
        if (theme["border-regular-control"]) {
            rules += ".action-btn, .text-input { border-color: " + theme["border-regular-control"] + "; }\n";
        }
        if (theme["highlight-button-hover"]) {
            rules += ".action-btn:hover { background-color: " + theme["highlight-button-hover"] + "; }\n";
        }
        if (theme["highlight-button-pressed"]) {
            rules += ".action-btn:active { background-color: " + theme["highlight-button-pressed"] + "; }\n";
        }
        if (theme["text-tertiary"]) {
            rules += ".section-title { color: " + theme["text-tertiary"] + "; }\n";
            rules += ".text-input::placeholder { color: " + theme["text-tertiary"] + "; }\n";
        }
        var styleEl = document.getElementById("pluginStyles");
        if (styleEl) styleEl.innerHTML = rules;
    };
})();
