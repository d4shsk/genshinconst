import json
import re
import os
import openpyxl # Убедитесь, что установили: python -m pip install openpyxl

def clean_text(text):
    """Убирает лишние пробелы и превращает None в пустую строку"""
    if text is None:
        return ""
    return str(text).strip()

def parse_genshin_xlsx(filename):
    if not os.path.exists(filename):
        print(f"❌ Файл {filename} не найден! Положите его в папку со скриптом.")
        return

    print(f"📂 Открываю Excel файл: {filename}...")
    
    # data_only=True важно, чтобы читать значения формул, а не сами формулы
    try:
        wb = openpyxl.load_workbook(filename, data_only=True)
    except Exception as e:
        print(f"❌ Ошибка при открытии файла: {e}")
        return

    # --- ВЫБОР ЛИСТА СОЗВЕЗДИЯ ---
    target_sheet_name = "СОЗВЕЗДИЯ"
    sheet = None

    if target_sheet_name in wb.sheetnames:
        print(f"✅ Нашел нужный лист: '{target_sheet_name}'")
        sheet = wb[target_sheet_name]
    else:
        print(f"⚠️ Лист '{target_sheet_name}' не найден!")
        print(f"   Доступные листы: {wb.sheetnames}")
        # Берем первый лист как запасной вариант
        sheet = wb.active 
        print(f"   👉 Использую первый лист: '{sheet.title}'")

    characters = []
    current_char = None
    
    # Регулярка для поиска С1..С6 (латиница C или кириллица С)
    constellation_regex = re.compile(r'^[СCcCсc]\s*([1-6])') 

    print("⚙️ Начинаю парсинг строк...")

    # Проходим по всем строкам таблицы
    for row in sheet.iter_rows():
        # Получаем список значений и список объектов ячеек (для комментариев)
        cells = list(row)
        values = [clean_text(cell.value) for cell in cells]
        
        # Если строка пустая, пропускаем
        if not any(values): continue

        # Ищем колонку с меткой "С1", "С2" и т.д.
        c_level = None
        c_idx = -1
        
        for idx, val in enumerate(values):
            match = constellation_regex.search(val)
            if match:
                c_level = "С" + match.group(1) # Нормализуем к русской С
                c_idx = idx
                break
        
        if c_level:
            # Данные: Урон, Поддержка, Описание идут СПРАВА от консты
            data_cells = cells[c_idx+1:]
            data_values = values[c_idx+1:]
            
            # Описание обычно последнее в блоке данных
            # Но нужно быть осторожным, иногда ячеек больше чем надо
            # Берем первые 3 значения после консты, если они есть
            
            damage = "-"
            support = "-"
            description = ""
            note_text = None

            # Эвристика на основе ваших данных:
            # Колонки: [УРОН, ПОДДЕРЖКА, БОНУС (ОПИСАНИЕ)]
            if len(data_values) >= 1: damage = data_values[0]
            if len(data_values) >= 2: support = data_values[1]
            if len(data_values) >= 3: description = data_values[2]

            # --- ПОИСК КОММЕНТАРИЯ (ПРИМЕЧАНИЯ) ---
            # Проверяем ячейки справа (Урон, Поддержка, Описание) на наличие заметки
            for cell in data_cells:
                if cell.comment:
                    note_text = cell.comment.text.strip()
                    # Часто Google Sheets добавляет имя автора в начале, можно почистить
                    # но пока оставим как есть
                    break

            # --- ЛОГИКА СОЗДАНИЯ ПЕРСОНАЖА ---
            if c_level == "С1":
                # Начинаем нового персонажа
                current_char = {
                    "name": "Unknown", 
                    "element": "?",
                    "constellations": {}
                }
                characters.append(current_char)
            
            if current_char:
                # ИМЯ ПЕРСОНАЖА ВСЕГДА НА СТРОКЕ С2 (согласно )
                if c_level == "С2":
                    # Имя обычно в самой первой колонке (индекс 0)
                    possible_name = values[0]
                    if possible_name and len(possible_name) > 1:
                        # Ищем эмодзи стихии
                        elem_match = re.search(r'([❄️🔥💧⚡️☘️💎💨])', possible_name)
                        element = elem_match.group(1) if elem_match else "?"
                        # Убираем стихию из имени
                        name_clean = possible_name.replace(element, "").strip()
                        
                        current_char["name"] = name_clean
                        current_char["element"] = element

                # Записываем данные в JSON
                current_char["constellations"][c_level] = {
                    "damage": damage,
                    "support": support,
                    "description": description,
                    "note": note_text  # Поле с комментарием
                }

    # Удаляем "пустых" или "сломанных" персонажей (где не нашли имя)
    final_chars = [c for c in characters if c["name"] != "Unknown" and c["name"] != "ПРИМЕР"]

    # Сохраняем
    with open('result.json', 'w', encoding='utf-8') as jf:
        json.dump(final_chars, jf, ensure_ascii=False, indent=2)
        print(f"🎉 Готово! Обработано персонажей: {len(final_chars)}")
        print("📁 Данные сохранены в 'result.json'")

# Запуск
# Убедитесь, что ваш файл называется data.xlsx (или поменяйте имя здесь)
parse_genshin_xlsx('data.xlsx')