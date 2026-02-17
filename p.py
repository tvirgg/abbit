import os
import zipfile
import xml.etree.ElementTree as ET

# Глобальный список исключаемых директорий
EXCLUDE_DIRS = ['node_modules', '.git', 'dist', '.next']
EXCEL_EXTENSIONS = ('.xlsx', '.xlsm', '.xltx', '.xltm')
EXCLUDE_FILES = {'.DS_Store'}

def is_binary_file(file_path, sample_size=1024):
    """
    Проверяет, является ли файл бинарным по наличию нулевых байтов в начале.
    """
    try:
        with open(file_path, 'rb') as file:
            chunk = file.read(sample_size)
            return b'\x00' in chunk
    except Exception:
        return True


def read_excel_file(file_path):
    """
    Возвращает текстовое представление Excel-файла (OOXML: xlsx/xltx/...).
    """
    ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    lines = []

    with zipfile.ZipFile(file_path, 'r') as zf:
        # Читаем sharedStrings (если есть)
        shared_strings = []
        if 'xl/sharedStrings.xml' in zf.namelist():
            root = ET.fromstring(zf.read('xl/sharedStrings.xml'))
            for si in root.findall('x:si', ns):
                chunks = [t.text or '' for t in si.findall('.//x:t', ns)]
                shared_strings.append(''.join(chunks))

        # Определяем имена листов
        sheet_names = {}
        if 'xl/workbook.xml' in zf.namelist():
            wb_root = ET.fromstring(zf.read('xl/workbook.xml'))
            rel_ns = {'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
            for sheet in wb_root.findall('.//x:sheets/x:sheet', ns):
                rel_id = sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                if rel_id:
                    sheet_names[rel_id] = sheet.attrib.get('name', rel_id)

            rels_path = 'xl/_rels/workbook.xml.rels'
            rel_map = {}
            if rels_path in zf.namelist():
                rel_root = ET.fromstring(zf.read(rels_path))
                for rel in rel_root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                    rid = rel.attrib.get('Id')
                    target = rel.attrib.get('Target', '')
                    if rid and target:
                        rel_map[rid] = 'xl/' + target.lstrip('/')

            ordered_sheets = []
            for rid, name in sheet_names.items():
                sheet_path = rel_map.get(rid)
                if sheet_path and sheet_path in zf.namelist():
                    ordered_sheets.append((name, sheet_path))
        else:
            ordered_sheets = []

        # fallback: если workbook.xml не разобрали, читаем все sheet*.xml
        if not ordered_sheets:
            ordered_sheets = []
            for name in sorted(zf.namelist()):
                if name.startswith('xl/worksheets/sheet') and name.endswith('.xml'):
                    ordered_sheets.append((os.path.basename(name), name))

        for sheet_name, sheet_path in ordered_sheets:
            lines.append(f'--- Sheet: {sheet_name} ---')
            sheet_root = ET.fromstring(zf.read(sheet_path))

            row_count = 0
            for row in sheet_root.findall('.//x:sheetData/x:row', ns):
                values = []
                for cell in row.findall('x:c', ns):
                    cell_type = cell.attrib.get('t')
                    v = cell.find('x:v', ns)
                    text = ''
                    if cell_type == 's' and v is not None and v.text is not None:
                        idx = int(v.text)
                        if 0 <= idx < len(shared_strings):
                            text = shared_strings[idx]
                    elif cell_type == 'inlineStr':
                        is_node = cell.find('x:is', ns)
                        if is_node is not None:
                            text_parts = [t.text or '' for t in is_node.findall('.//x:t', ns)]
                            text = ''.join(text_parts)
                    elif v is not None and v.text is not None:
                        text = v.text
                    values.append(text)

                if any(val != '' for val in values):
                    lines.append('\t'.join(values))
                    row_count += 1

            if row_count == 0:
                lines.append('[Лист пустой]')
            lines.append('')

    return '\n'.join(lines).strip()


def generate_directory_tree(start_path, output_file, output_filename, exclude_dirs=None):
    """
    Генерирует дерево каталогов и пишет его в выходной файл.
    """
    if exclude_dirs is None:
        exclude_dirs = EXCLUDE_DIRS

    output_file.write("# Project Directory Structure\n\n")

    for root, dirs, files in os.walk(start_path):
        # Исключаем ненужные папки
        dirs[:] = [d for d in dirs if d not in exclude_dirs]

        # Уровень вложенности
        level = root.replace(start_path, '').count(os.sep)
        indent = '  ' * level

        # Относительный путь
        rel_path = os.path.relpath(root, start_path)
        if rel_path == '.':
            output_file.write(f"{indent}📁 ./\n")
        else:
            dir_name = os.path.basename(root)
            output_file.write(f"{indent}📁 {dir_name}/\n")

        # Файлы
        sub_indent = '  ' * (level + 1)
        for file in sorted(files):
            if file == output_filename or file in EXCLUDE_FILES:
                continue
            output_file.write(f"{sub_indent}📄 {file}\n")

    output_file.write("\n\n# File Contents\n\n")


def main():
    # Текущая директория
    current_directory = os.getcwd()

    # Имя выходного файла
    output_filename = 'files_with_code.txt'

    with open(output_filename, 'w', encoding='utf-8') as output:
        # Генерируем дерево каталогов
        generate_directory_tree(current_directory, output, output_filename, EXCLUDE_DIRS)

        # Записываем содержимое файлов
        for root, dirs, files in os.walk(current_directory):
            # Исключаем ненужные папки
            dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS]

            for filename in files:
                if filename == output_filename or filename in EXCLUDE_FILES:
                    continue

                # Относительный путь
                rel_path = os.path.relpath(os.path.join(root, filename), current_directory)
                file_path = os.path.join(root, filename)

                # Заголовок файла
                output.write(f'File: {rel_path}\n')
                output.write('=' * 80 + '\n')

                try:
                    if filename.lower().endswith(EXCEL_EXTENSIONS):
                        excel_content = read_excel_file(file_path)
                        output.write(excel_content + '\n\n')
                        output.write('-' * 80 + '\n\n')
                        continue

                    if is_binary_file(file_path):
                        size = os.path.getsize(file_path)
                        output.write(f'[Бинарный файл, размер: {size} байт]\n\n')
                        output.write('-' * 80 + '\n\n')
                        continue

                    with open(file_path, 'r', encoding='utf-8') as file:
                        code = file.read()
                        output.write(code + '\n\n')
                        output.write('-' * 80 + '\n\n')
                except Exception as e:
                    output.write(f'Не удалось прочитать файл {rel_path}: {e}\n\n')
                    output.write('-' * 80 + '\n\n')

    print(f'Дерево каталогов и код файлов записаны в {output_filename}')


if __name__ == "__main__":
    main()
