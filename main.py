import pandas as pd
from fuzzywuzzy import process, fuzz
from cleantext import clean
import re

df_KA = pd.read_excel('1.xlsx')
df_GAP = pd.read_excel('2.xlsx')

# Функция для очистки названия компании
def clean_company_name(company_name):
    cleaned_name = clean(company_name,
                    clean_all= True, # Execute all cleaning operations
                    extra_spaces=True ,  # Remove extra white spaces 
                    stemming=True , # Stem the words
                    stopwords=True ,# Remove stop words
                    lowercase=True ,# Convert to lowercase
                    numbers=True ,# Remove all digits 
                    punct=True ,# Remove all punctuations
                    reg=r'\b(ООО|ОАО|АО|ЗАО|ПФ|ПАО|L.L.C|ИП|ТОО|Ltd|Co.)\b', # Remove parts of text based on regex
                    #reg_replace: str = '<replace_value>', # String to replace the regex used in reg
                    #stp_lang='english'  # Language for stop words
                    )
    # Удаляем специфические кавычки
    cleaned_name = re.sub(r'[«»]', '', cleaned_name)
    
    return cleaned_name

# Создаем словарь для хранения результатов
my_dict = {}

# Очищаем названия компаний из обеих DataFrame
df_GAP['cleaned'] = df_GAP['Компания'].apply(clean_company_name)
df_KA['cleaned'] = df_KA['Контрагент'].apply(clean_company_name)

# Используем process.extract для нахождения совпадений
for num, cleaned_company in df_GAP['cleaned'].items():
    matches = process.extract(cleaned_company, df_KA['cleaned'], limit=None, scorer=fuzz.ratio)
    my_list = [df_KA['Контрагент'][_] for i, score, _ in matches if -

    if my_list:  # добавляем только если есть совпадения
        #my_dict[cleaned_company] = my_list
        my_dict[df_GAP.loc[num, "Компания"]] = my_list

# Печатаем результаты
for cleaned_name, original_names in my_dict.items():
    print(cleaned_name)
    for original_name in original_names:
        print(f'            {original_name}')
    print()
