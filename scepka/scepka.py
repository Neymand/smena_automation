def convert_uuids(input_file, output_file):
    """
    Читает UUID из входного файла и записывает их в выходной файл
    в формате с кавычками и запятыми
    """
    try:
        # Читаем данные из входного файла
        with open(input_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        # Обрабатываем каждую строку
        processed_lines = []
        for line in lines:
            line = line.strip()  # Убираем лишние пробелы и переносы строк
            if line:  # Если строка не пустая
                processed_lines.append(f"'{line}'")

        # Записываем результат в выходной файл
        with open(output_file, 'w', encoding='utf-8') as f:
            # Добавляем запятые после каждой строки, кроме последней
            for i, line in enumerate(processed_lines):
                if i < len(processed_lines) - 1:
                    f.write(f"{line},")
                else:
                    f.write(line)

        print(f"Преобразование завершено! Результат записан в файл: {output_file}")

    except FileNotFoundError:
        print(f"Ошибка: Файл {input_file} не найден")
    except Exception as e:
        print(f"Произошла ошибка: {e}")


# Использование скрипта
if __name__ == "__main__":
    input_filename = "text.txt"  # Имя входного файла
    output_filename = "output.txt"  # Имя выходного файла

    convert_uuids(input_filename, output_filename)