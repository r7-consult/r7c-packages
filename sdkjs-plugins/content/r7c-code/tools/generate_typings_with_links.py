import os
import re
from collections import defaultdict

EXAMPLES_DIR = os.path.join('modules', 'macros_ide', 'resources', 'examples')
OUTPUT_PATH = os.path.join('modules', 'macros_ide', 'ui', 'typings', 'onlyoffice-api.d.ts')

CALL_PATTERN = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\s*\.\s*([A-Za-z_][A-Za-z0-9_]*)\s*\(")

ALIAS_MAP = {
    'Api': 'Api',
    'ApiWorksheet': 'ApiWorksheet',
    'ApiRange': 'ApiRange',
    'ApiWorkbook': 'ApiWorkbook',
    'sheet': 'ApiWorksheet',
    'Sheet': 'ApiWorksheet',
    'range': 'ApiRange',
    'Range': 'ApiRange',
    'workbook': 'ApiWorkbook',
    'wb': 'ApiWorkbook',
}


def normalise_object(folder: str, obj: str) -> str:
    if folder.startswith('Api'):
        return folder
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


def generate():
    method_map = defaultdict(lambda: defaultdict(set))

    for root, _, files in os.walk(EXAMPLES_DIR):
        folder = os.path.basename(root)
        for fname in files:
            if not fname.lower().endswith(('.js', '.md')):
                continue
            path = os.path.join(root, fname)
            relative = os.path.relpath(path, EXAMPLES_DIR)
            try:
                with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
            except OSError:
                continue
            for obj, method in re.findall(r"\b([A-Za-z_][A-Za-z0-9_]*)\s*\.\s*([A-Za-z_][A-Za-z0-9_]*)\s*\(", content):
                key = normalise_object(folder, obj)
                method_map[key][method].add(relative)

    method_map.setdefault('Api', {})

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    lines = [
        '/**',
        ' * Auto-generated OnlyOffice macro typings with example links.',
        f' * Source: {EXAMPLES_DIR}/**/*.',
        ' */',
        'declare namespace OnlyOffice {'
    ]

    for key in sorted(method_map):
        methods = method_map[key]
        if not methods:
            continue
        iface = interface_name(key)
        lines.append(f'  interface {iface} {{')
        for method in sorted(methods):
            link_entries = []
            for rel in sorted(methods[method]):
                uri = os.path.join('modules/macros_ide/resources/examples', rel).replace(os.sep, '/')
                link_entries.append(f'{{@link {uri}}}')
            if link_entries:
                lines.append(f'    /** Example: {" ".join(link_entries)} */')
            lines.append(f'    {method}(...args: any[]): any;')
        lines.append('  }\n')

    lines.append('}')
    lines.append('')
    lines.append('declare const Api: OnlyOffice.DocumentApi;')
    lines.append('// ambient declarations for Monaco')

    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))


if __name__ == '__main__':
    generate()
