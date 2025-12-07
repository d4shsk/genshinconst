import json
import re
import os
import openpyxl

def format_value(val):
    if val is None: return ""
    # Превращаем 0.18 в 18%
    if isinstance(val, (float, int)):
        if 0 < abs(val) < 3.0: 
            return f"{val * 100:.1f}%".replace(".0%", "%")
        return str(val)
    return str(val).strip()

def parse_genshin_xlsx(filename):
    if not os.path.exists(filename):
        print(f"❌ Файл {filename} не найден!")
        return

    print(f"📂 Читаю файл: {filename}...")
    try:
        wb = openpyxl.load_workbook(filename, data_only=True)
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        return

    # Ищем лист
    target_name = "СОЗВЕЗДИЯ"
    if target_name in wb.sheetnames:
        sheet = wb[target_name]
    else:
        sheet = wb.active
        print(f"⚠️ Лист '{target_name}' не найден, беру первый попавшийся.")

    characters = []
    current_char = None
    constellation_regex = re.compile(r'^[СCcCсc]\s*([1-6])') 

    # Регулярка для энергии: Число + Е (лат/кир), например "15E", "7,2Е"
    energy_regex = re.compile(r'^[\d.,]+\s*[EЕеe]$')

    for row in sheet.iter_rows():
        cells = list(row)
        values = [format_value(cell.value) for cell in cells]
        if not any(values): continue

        c_level = None
        c_idx = -1
        
        for idx, val in enumerate(values):
            match = constellation_regex.search(val)
            if match:
                c_level = "С" + match.group(1)
                c_idx = idx
                break
        
        if c_level:
            data_values = values[c_idx+1:]
            data_cells = cells[c_idx+1:]

            damage = data_values[0] if len(data_values) >= 1 else "-"
            support = data_values[1] if len(data_values) >= 2 else "-"
            description = data_values[2] if len(data_values) >= 3 else ""
            
            # --- ЛОГИКА ЭНЕРГИИ ---
            energy_val = None

            # Если в Уроне написано "15E", переносим в энергию
            if energy_regex.match(damage):
                energy_val = damage
                damage = "-" # Убираем из урона, чтобы не портило сортировку
            
            # Если вдруг в Поддержке написано "15E"
            elif energy_regex.match(support):
                energy_val = support
                support = "-"

            # --- Поиск комментария ---
            note_text = None
            for cell in data_cells:
                if cell.comment:
                    note_text = cell.comment.text.strip()
                    break

            # Создание персонажа
            if c_level == "С1":
                current_char = { "name": "Unknown", "element": "?", "constellations": {} }
                characters.append(current_char)
            
            if current_char:
                if c_level == "С2":
                    possible_name = values[0]
                    if possible_name and len(possible_name) > 1 and "ПРИМЕР" not in possible_name:
                        elem_match = re.search(r'([❄️🔥💧⚡️☘️💎💨])', possible_name)
                        element = elem_match.group(1) if elem_match else "?"
                        name_clean = possible_name.replace(element, "").strip()
                        current_char["name"] = name_clean
                        current_char["element"] = element

                current_char["constellations"][c_level] = {
                    "damage": damage,
                    "support": support,
                    "description": description,
                    "note": note_text,
                    "energy": energy_val # Новое поле
                }

    final_chars = [c for c in characters if c["name"] != "Unknown" and "ПРИМЕР" not in c["name"].upper()]

    with open('result.json', 'w', encoding='utf-8') as jf:
        json.dump(final_chars, jf, ensure_ascii=False, indent=2)
        print(f"✅ Готово! Сохранено персонажей: {len(final_chars)}")

if __name__ == "__main__":
    parse_genshin_xlsx('data.xlsx')