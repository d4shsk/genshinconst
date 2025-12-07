import csv
import json
import re
import os

def clean_text(text):
    """Убирает лишние пробелы"""
    if text:
        return text.strip()
    return ""

def parse_genshin_file(filename):
    characters = []
    current_char = None
    
    # Регулярка для поиска С1..С6 (латиница C или кириллица С)
    constellation_regex = re.compile(r'^[СCcC]([1-6])') 

    # Проверка, существует ли файл
    if not os.path.exists(filename):
        return json.dumps({"error": f"Файл {filename} не найден!"}, ensure_ascii=False)

    # Открываем файл. encoding='utf-8' важен для русских букв и эмодзи
    with open(filename, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        
        for parts in reader:
            # parts - это уже список строк, кавычки csv.reader убрал сам
            if not parts: continue

            # Чистим пробелы в каждом элементе
            parts = [clean_text(p) for p in parts]

            # Ищем колонку с "С1", "С2"...
            c_level = None
            c_idx = -1
            
            for idx, p in enumerate(parts):
                match = constellation_regex.search(p)
                if match:
                    c_level = "С" + match.group(1) # Нормализуем к русской С
                    c_idx = idx
                    break
            
            if c_level:
                # Данные: Урон, Поддержка, Описание
                data_parts = parts[c_idx+1:]
                
                description = data_parts[-1] if data_parts else ""
                # Если всего 2 поля после Консты, значит первое - Урон, второе - Описание (поддержки нет)
                # Если 3 поля - Урон, Поддержка, Описание
                
                # Попробуем найти урон и поддержку, отталкиваясь от конца списка
                # Обычно описание - последнее. Остальное между Сх и Описанием.
                middle_stats = data_parts[:-1]
                
                damage = "-"
                support = "-"

                if len(middle_stats) >= 1:
                    damage = middle_stats[0]
                if len(middle_stats) >= 2:
                    support = middle_stats[1]

                # --- ЛОГИКА СОЗДАНИЯ ПЕРСОНАЖА ---
                if c_level == "С1":
                    # Новый блок персонажа
                    current_char = {
                        "name": "Unknown", 
                        "element": "?",
                        "constellations": {}
                    }
                    characters.append(current_char)
                
                if current_char:
                    # Имя всегда в 1-й колонке (индекс 0) на строке С2
                    if c_level == "С2":
                        possible_name = parts[0]
                        if possible_name and len(possible_name) > 1:
                            elem_match = re.search(r'([❄️🔥💧⚡️☘️💎💨])', possible_name)
                            element = elem_match.group(1) if elem_match else "?"
                            name_clean = possible_name.replace(element, "").strip()
                            current_char["name"] = name_clean
                            current_char["element"] = element

                    # Записываем данные созвездия
                    current_char["constellations"][c_level] = {
                        "damage": damage,
                        "support": support,
                        "description": description
                    }

    # Сохраняем результат в json файл (опционально)
    with open('result.json', 'w', encoding='utf-8') as json_file:
        json.dump(characters, json_file, ensure_ascii=False, indent=2)
        print("Готово! Данные сохранены в result.json")

    return json.dumps(characters, ensure_ascii=False, indent=2)

# Запуск
# Убедитесь, что ваш csv файл называется data.csv
print(parse_genshin_file('data.csv'))