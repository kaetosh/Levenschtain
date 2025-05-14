import pandas as pd
from fuzzywuzzy import process, fuzz
from cleantext import clean
import re
from typing import List, Tuple

# Чтение данных из Excel
df_KA: pd.DataFrame = pd.read_excel('1.xlsx') # список контрагентов
df_GAP: pd.DataFrame = pd.read_excel('2.xlsx') # список компаний группы

# Функция для очистки названия компании
def clean_company_name(company_name: str) -> str:
    cleaned_name: str = clean(company_name,
                    clean_all=True,  # Выполняем все операции очистки
                    extra_spaces=True,  # Удаляем лишние пробелы 
                    stemming=True,  # Стеммим слова
                    stopwords=True,  # Удаляем стоп-слова
                    stp_lang='russian', # Язык стоп-слов
                    lowercase=True,  # Приводим к нижнему регистру
                    #numbers=True,  # Удаляем все цифры
                    punct=True,  # Удаляем все знаки препинания
                    reg=r'\b(ООО|ОАО|АО|ЗАО|ПФ|ПАО|L.L.C|ИП|ТОО|Ltd|Co.НП|СО|КП|ФК|ГК|ЗАО|ОАО|ПАО|ИП|ТОО|LLP|PLC|S.A.|S.R.L.|GmbH|B.V.|Inc.|Corp.|S.p.A.|Pty Ltd|SAS|N.V.)\b',  # Удаляем части текста по regex
                    )
    # Удаляем специфические кавычки
    cleaned_name = re.sub(r'[«»]', '', cleaned_name)
    return cleaned_name

# Очищаем названия компаний из обеих DataFrame
df_GAP['cleaned'] = df_GAP['Компания'].apply(clean_company_name)
df_KA['cleaned'] = df_KA['Контрагент'].apply(clean_company_name)

# Функция для нахождения совпадений
def find_matches(cleaned_company: str) -> List[str]:
    matches: List[Tuple[str, int, int]] = process.extract(cleaned_company, df_KA['cleaned'], limit=None, scorer=fuzz.ratio)
    result: List[str] = [df_KA['Контрагент'][ind] for _, score, ind in matches if score > 90]
    return result 

# Применяем функцию для нахождения совпадений
df_GAP['matches'] = df_GAP['cleaned'].apply(find_matches)

df_GAP = df_GAP[df_GAP['matches'].apply(lambda x: x != [])]
df_GAP = df_GAP.explode('matches')
df_GAP[['matches','Компания']].to_excel('3.xlsx', index=False)
