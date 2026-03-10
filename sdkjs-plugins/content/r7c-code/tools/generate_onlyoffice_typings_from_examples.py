import os
import re
from collections import defaultdict

EXAMPLES_DIR = os.path.join(
    "tmp",
    "office-js-api-master",
    "Examples"
)
OUTPUT_PATH = os.path.join(
    "modules",
    "macros_ide",
    "ui",
    "typings",
    "onlyoffice-api.d.ts"
)

# pattern to catch object.method(
CALL_PATTERN = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\s*\.\s*([A-Za-z_][A-Za-z0-9_]*)\s*\(")

# heuristics to map variable names to API object names
ALIAS_MAP = {
    "Api": "Api",
    "ApiRange": "ApiRange",
    "ApiWorksheet": "ApiWorksheet",
    "ApiWorkbook": "ApiWorkbook",
    "sheet": "ApiWorksheet",
    "Sheet": "ApiWorksheet",
    "worksheet": "ApiWorksheet",
    "ws": "ApiWorksheet",
    "range": "ApiRange",
    "Range": "ApiRange",
    "cell": "ApiRange",
    "workbook": "ApiWorkbook",
    "wb": "ApiWorkbook",
}


def normalise_object(folder: str, obj: str) -> str:
    if folder.startswith("Api"):
        return folder
    if obj in ALIAS_MAP:
        return ALIAS_MAP[obj]
    simple = obj.lstrip("o") or obj
    return "Api" + simple[0].upper() + simple[1:]


def interface_name(key: str) -> str:
    if key == "Api":
        return "DocumentApi"
    if key.startswith("Api"):
        return key[3:] or "Unnamed"
    return key[0].upper() + key[1:]


def generate():
    method_map = defaultdict(set)

    for root, _, files in os.walk(EXAMPLES_DIR):
        folder = os.path.basename(root)
        for fname in files:
            if not fname.lower().endswith(".js"):
                continue
            path = os.path.join(root, fname)
            try:
                with open(path, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()
            except OSError:
                continue
            for obj, method in CALL_PATTERN.findall(content):
                key = normalise_object(folder, obj)
                method_map[key].add(method)

    method_map.setdefault("Api", set())

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)

    lines = [
        "/**",
        " * Auto-generated OnlyOffice macro typings from JS examples.",
        f" * Source: {EXAMPLES_DIR}/**/*.js",
        " */",
        "declare namespace OnlyOffice {",
    ]

    for key in sorted(method_map):
        methods = sorted(method_map[key])
        if not methods:
            continue
        iface = interface_name(key)
        lines.append(f"  interface {iface} {{")
        for method in methods:
            lines.append(f"    {method}(...args: any[]): any;")
        lines.append("  }\n")

    lines.append("}")
    lines.append("")
    lines.append("declare const Api: OnlyOffice.DocumentApi;")
    lines.append("// ambient declarations for Monaco")

    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


if __name__ == "__main__":
    generate()
