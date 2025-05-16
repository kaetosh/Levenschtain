import pandas as pd
from fuzzywuzzy import process, fuzz
from cleantext import clean
import re
from typing import List, Tuple
from data_text import NAME_DATA_FILE, NAME_OUTPUT_FILE



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
                    reg=r'\b(ООО|ОАО|АО|ЗАО|ПФ|ПАО|L.L.C|ИП|ТОО|Ltd|Co.НП|СО|КП|ФК|ГК|ЗАО|ОАО|ПАО|ИП|ТОО|LLP|PLC|S.A.|S.R.L.|GmbH|B.V.|Inc.|Corp.|S.p.A.|Pty Ltd|SAS|N.V.)\b'  # Удаляем части текста по regex
                    )
    # Удаляем специфические кавычки
    cleaned_name = re.sub(r'[«»]', '', cleaned_name)
    return cleaned_name

# Функция для нахождения совпадений
def find_matches(cleaned_company: str, df_data_b: pd.DataFrame) -> List[str]:
    matches: List[Tuple[str, int, int]] = process.extract(cleaned_company, df_data_b['cleaned'], limit=None, scorer=fuzz.ratio)
    result: List[str] = [df_data_b['data2'][ind] for _, score, ind in matches if score > 90]
    return result

# Основная функция для создания файла совпадений
def create_file_matches() -> None:
    df_data: pd.DataFrame = pd.read_excel(NAME_DATA_FILE)
    df_data_a = pd.DataFrame(df_data['data1'].dropna().astype(str))
    df_data_b = pd.DataFrame(df_data['data2'].dropna().astype(str))


    df_data_a['cleaned'] = df_data_a['data1'].apply(clean_company_name)
    df_data_b['cleaned'] = df_data_b['data2'].apply(clean_company_name)

    df_output = pd.DataFrame(df_data_a['data1'])

    df_output.to_excel('11.xlsx')

    df_output['Совпадения'] = df_data_a['cleaned'].apply(find_matches, df_data_b=df_data_b)
    df_output = df_output[df_output['Совпадения'].apply(lambda x: x != [])]
    df_output = df_output.explode('Совпадения')
    df_output[['data1', 'Совпадения']].to_excel(NAME_OUTPUT_FILE, index=False)
