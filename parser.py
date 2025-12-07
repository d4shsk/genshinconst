import json
import re
import os
import openpyxl

def format_value(val):
    if val is None: return ""
    if isinstance(val, (float, int)):
        if 0 < abs(val) < 3.0: 
            return f"{val * 100:.1f}%".replace(".0%", "%")
        return str(val)
    return str(val).strip()

def clean_damage_value(val):
    """Превращает '576' или '2,7' в число 576.0"""
    if not val or str(val).strip() == "-": return 0
    # Убираем пробелы, меняем запятую на точку
    clean = str(val).replace(",", ".").replace(" ", "").strip()
    try:
        return float(clean)
    except:
        return 0

def parse_genshin_xlsx(filename):
    if not os.path.exists(filename):
        print(f"❌ Файл {filename} не найден!")
        return

    print(f"📂 Открываю Excel файл: {filename}...")
    try:
        wb = openpyxl.load_workbook(filename, data_only=True)
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        return

    # --- 1. ПАРСИНГ ЛИСТА "ЭФФЕКТИВНОСТЬ" ---
    stats_map = {} 
    
    sheet_eff_name = "ЭФФЕКТИВНОСТЬ"
    if sheet_eff_name in wb.sheetnames:
        print(f"✅ Читаю лист '{sheet_eff_name}'...")
        sheet_eff = wb[sheet_eff_name]
        
        # Индексы колонок (A=0 ... O=14)
        # I(8)=Урон на поле, K(10)=УНП Макс
        # N(13)=Урон с кармана, O(14)=УСК Макс
        
        for row in sheet_eff.iter_rows(min_row=2):
            cells = [c.value for c in row]
            # Нам нужно как минимум 15 колонок (до O включительно)
            if len(cells) < 15: continue 
            
            rarity_raw = str(cells[0]).strip() 
            name_raw = str(cells[1]).strip()   
            
            if not name_raw or name_raw == "None": continue
            
            # Считываем урон
            dmg_field = clean_damage_value(cells[8])       # I
            dmg_field_max = clean_damage_value(cells[10])  # K
            dmg_off_field = clean_damage_value(cells[13])  # N
            dmg_off_field_max = clean_damage_value(cells[14]) # O (Новое)
            
            rarity = "5" if "5" in rarity_raw else "4"
            
            stats_map[name_raw] = {
                "rarity": rarity,
                "base_stats": {
                    "field": dmg_field,
                    "field_max": dmg_field_max,
                    "off_field": dmg_off_field,
                    "off_field_max": dmg_off_field_max
                }
            }
    else:
        print(f"⚠️ Лист '{sheet_eff_name}' не найден! Базовый урон будет равен 0.")

    # --- 2. ПАРСИНГ ЛИСТА "СОЗВЕЗДИЯ" ---
    sheet_const = wb["СОЗВЕЗДИЯ"] if "СОЗВЕЗДИЯ" in wb.sheetnames else wb.active
    print(f"✅ Читаю лист '{sheet_const.title}'...")

    characters = []
    current_char = None
    constellation_regex = re.compile(r'^[СCcCсc]\s*([1-6])') 
    energy_regex = re.compile(r'^[\d.,]+\s*[EЕеe]$')

    for row in sheet_const.iter_rows():
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
            
            energy_val = None
            if energy_regex.match(damage):
                energy_val = damage
                damage = "-"
            elif energy_regex.match(support):
                energy_val = support
                support = "-"

            note_text = None
            for cell in data_cells:
                if cell.comment:
                    note_text = cell.comment.text.strip()
                    break

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

                        # Привязка данных
                        found_stats = stats_map.get(name_clean)
                        if not found_stats:
                            for k, v in stats_map.items():
                                k_norm = k.replace("(", "").replace(")", "").replace("-", "").replace(" ", "").lower()
                                name_norm = name_clean.replace("(", "").replace(")", "").replace("-", "").replace(" ", "").lower()
                                if name_norm in k_norm or k_norm in name_norm:
                                    found_stats = v
                                    break
                        
                        if found_stats:
                            current_char["rarity"] = found_stats["rarity"]
                            current_char["base_stats"] = found_stats["base_stats"]
                        else:
                            current_char["rarity"] = "?" 
                            current_char["base_stats"] = {"field": 0, "field_max": 0, "off_field": 0, "off_field_max": 0}

                current_char["constellations"][c_level] = {
                    "damage": damage,
                    "support": support,
                    "description": description,
                    "note": note_text,
                    "energy": energy_val
                }

    final_chars = [c for c in characters if c["name"] != "Unknown" and "ПРИМЕР" not in c["name"].upper()]

    with open('result.json', 'w', encoding='utf-8') as jf:
        json.dump(final_chars, jf, ensure_ascii=False, indent=2)
        print(f"✅ Готово! Сохранено: {len(final_chars)}. Данные обновлены (4 колонки урона).")

if __name__ == "__main__":
    parse_genshin_xlsx('data.xlsx')