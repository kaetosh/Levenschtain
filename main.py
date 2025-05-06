import pandas as pd
from fuzzywuzzy import process, fuzz
from cleantext import clean
import re
from typing import Dict, List

# Чтение данных из Excel
df_KA: pd.DataFrame = pd.read_excel('1.xlsx')
df_GAP: pd.DataFrame = pd.read_excel('2.xlsx')

# Функция для очистки названия компании
def clean_company_name(company_name: str) -> str:
    cleaned_name: str = clean(company_name,
                    clean_all=True,  # Выполняем все операции очистки
                    extra_spaces=True,  # Удаляем лишние пробелы 
                    stemming=True,  # Стеммим слова
                    stopwords=True,  # Удаляем стоп-слова
                    lowercase=True,  # Приводим к нижнему регистру
                    numbers=True,  # Удаляем все цифры 
                    punct=True,  # Удаляем все знаки препинания
                    reg=r'\b(ООО|ОАО|АО|ЗАО|ПФ|ПАО|L.L.C|ИП|ТОО|Ltd|Co.)\b',  # Удаляем части текста по regex
                    )
    # Удаляем специфические кавычки
    cleaned_name = re.sub(r'[«»]', '', cleaned_name)
    return cleaned_name

# Создаем словарь для хранения результатов
my_dict: Dict[str, List[str]] = {}

# Очищаем названия компаний из обеих DataFrame
df_GAP['cleaned'] = df_GAP['Компания'].apply(clean_company_name)
df_KA['cleaned'] = df_KA['Контрагент'].apply(clean_company_name)

# Функция для нахождения совпадений
def find_matches(cleaned_company: str) -> List[str]:
    matches: List[Tuple[str, int, int]] = process.extract(cleaned_company, df_KA['cleaned'], limit=None, scorer=fuzz.ratio)
    result: List[str] = [df_KA['Контрагент'][index] for _, score, index in matches if score > 90]
    return result 

# Применяем функцию для нахождения совпадений
df_GAP['matches'] = df_GAP['cleaned'].apply(find_matches)

# Заполняем словарь результатами
for index, row in df_GAP.iterrows():
    if row['matches']:  # добавляем только если есть совпадения
        my_dict[row["Компания"]] = row['matches']

df_GAP = df_GAP[df_GAP['matches'].apply(lambda x: x != [])]
df_GAP.loc[:, 'matches'] = df_GAP['matches'].apply(lambda x: ', '.join(x))
df_GAP[['Компания', 'matches']].to_excel('3.xlsx', index=False)
