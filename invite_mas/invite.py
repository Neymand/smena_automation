import json

# Исходный JSON документ (шаблон)
template = {
    "group_id": "group_id",
    "allow_create_shifts": True,
    "allow_changing_shift_rate": False
}

# Список group_id
group_ids = [
]
count = 0
with open("list_tt.txt", "r", encoding="utf-8") as f:
    group_ids.extend(map(str.strip, f))

#Создаем список объектов
result = []

for group_id in group_ids:
    # Создаем копию шаблона и заменяем group_id
    item = template.copy()
    item["group_id"] = group_id
    result.append(item)
    count += 1

# Сохраняем в файл
with open("output.json", "w", encoding="utf-8") as f:
    json.dump(result, f, indent=2, ensure_ascii=False)

print("Файл output.json успешно создан!")

print("Инвайт ссылок обработано", count)