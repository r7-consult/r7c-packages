import os
import re
from collections import defaultdict

BASE_DIR = os.path.join('modules', 'macros_ide', 'resources', 'examples')
OUTPUT_PATH = os.path.join('modules', 'macros_ide', 'ui', 'typings', 'onlyoffice-api.d.ts')
CODE_BLOCK_PATTERN = re.compile(r"```[a-zA-Z]*\n(.*?)```", re.DOTALL)

METHOD_ALIAS = {
    "GetRange": "GetRange",
    "GetCells": "GetRange",
    "GetCell": "GetRange",
}
ALIAS_MAP = {
    "sheet": "ApiWorksheet",
    "Sheet": "ApiWorksheet",
    "range": "ApiRange",
    "Range": "ApiRange",
}


def extract_calls(code: str):
    for match in re.finditer(r"\b([A-Za-z_][A-Za-z0-9_]*)\s*\.\s*([A-Za-z_][A-Za-z0-9_]*)\s*\(", code):
        yield match.group(1), match.group(2)


def guess_key(folder: str, obj: str, method: str) -> str:
    if folder.startswith('Api'):
        return folder
    if obj == 'Api':
        return 'Api'
    if obj in ALIAS_MAP:
        return ALIAS_MAP[obj]
    simple = obj.lstrip('o') or obj
    return 'Api' + simple[0].upper() + simple[1:]


def interface_name(key: str) -> str:
    if key == 'Api':
        return 'DocumentApi'
    if key.startswith('Api'):
        return key[3:] or 'Unnamed'
    return key[0].upper() + key[1:]


def normalise_method(method: str) -> str:
    return METHOD_ALIAS.get(method, method)


def generate():
    method_map = defaultdict(set)

    for root, _, files in os.walk(BASE_DIR):
        folder = os.path.basename(root)
        for fname in files:
            if not fname.lower().endswith('.md'):
                continue
            path = os.path.join(root, fname)
            with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            for block in CODE_BLOCK_PATTERN.findall(content):
                for obj, method in extract_calls(block):
                    key = guess_key(folder, obj, method)
                    method_map[key].add(normalise_method(method))

    method_map.setdefault('Api', set())

    lines = [
        '/**',
        ' * Auto-generated OnlyOffice macro typings (simplified).',
        f' * Source: {BASE_DIR}/*.md',
        ' */',
        'declare namespace OnlyOffice {'
    ]

    for key in sorted(method_map):
        methods = sorted(method_map[key])
        if not methods:
            continue
        iface = interface_name(key)
        lines.append(f'  interface {iface} {{')
        for method in methods:
            lines.append(f'    {method}(...args: any[]): any;')
        lines.append('  }\n')

    lines.append('}')
    lines.append('')
    lines.append('declare const Api: OnlyOffice.DocumentApi;')
    lines.append('// ambient declarations for Monaco')

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))


if __name__ == '__main__':
    generate()
