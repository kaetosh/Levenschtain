import re
from difflib import SequenceMatcher

def normalize_string(s):
    # Приведение к нижнему регистру
    s = s.lower()
    # Удаление специальных символов и лишних пробелов
    s = re.sub(r'\b(ооо|г\.|и др\.|и т.д.)\b', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def similarity(s1, s2):
    return SequenceMatcher(None, s1, s2).ratio()

def find_similar_companies(group_companies, contractors, threshold=0.8):
    normalized_group = [normalize_string(company) for company in group_companies]
    normalized_contractors = [normalize_string(contractor) for contractor in contractors]

    matches = {}
    for contractor in normalized_contractors:
        for group_company in normalized_group:
            if similarity(contractor, group_company) >= threshold:
                matches[contractor] = group_company
                break  # Выходим из внутреннего цикла, если нашли совпадение

    return matches

# Пример использования
group_companies = ["ООО Северо-Кавказский Бройлер"]
contractors = ["Северо кавказский бройлео ООО", "СевероКавказский Бройлер"]

matches = find_similar_companies(group_companies, contractors)
print(matches)
