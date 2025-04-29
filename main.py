pip install fuzzywuzzy
pip install python-Levenshtein  # Опционально, для ускорения
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

str1 = "apple"
str2 = "appl"

# Сравнение двух строк
similarity = fuzz.ratio(str1, str2)
print(f"Сходство: {similarity}%")

# Поиск наиболее похожей строки из списка
choices = ["apple", "banana", "grape"]
best_match = process.extractOne("appl", choices)
print(f"Лучший вариант: {best_match}")
