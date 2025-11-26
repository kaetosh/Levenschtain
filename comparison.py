import os
import pandas as pd
from rapidfuzz import process, fuzz, utils
from typing import Dict
from text import NAME_DATA_FILE, NAME_OUTPUT_FILE
from collections import defaultdict

from custom_errors import Sheet_too_large_Error
from utils import load_config, clean_text_optimized # Импортируем наши новые функции

# Кэширование результатов очистки
_cleaning_cache: Dict[str, str] = {}

def clean_company_name(company_name: str) -> str:
    """
    Оптимизированная функция очистки названия компании с использованием 
    опциональных параметров из конфигурационного файла.
    """
    if company_name in _cleaning_cache:
        return _cleaning_cache[company_name]
    
    # Загружаем конфигурацию
    config = load_config()
    
    # Используем оптимизированную функцию очистки
    normalized_words: str = clean_text_optimized(company_name, config)
    
    _cleaning_cache[company_name] = normalized_words
    return normalized_words

def create_file_matches(similarity_criterion: int) -> None:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(script_dir, "working_files", NAME_DATA_FILE)
    
    df_data: pd.DataFrame = pd.read_excel(file_path)
    # Оптимизация: работа с данными без копирования
    df_data_a = df_data[['data1']].dropna().drop_duplicates()
    df_data_b = df_data[['data2']].dropna().drop_duplicates()
    
    # Векторизованная очистка данных
    df_data_a = df_data_a.copy()
    df_data_b = df_data_b.copy()
    
    # Очистка данных с использованием новой функции
    df_data_a['cleaned'] = df_data_a['data1'].apply(clean_company_name)
    df_data_b['cleaned'] = df_data_b['data2'].apply(clean_company_name)
    
    # Создаем словарь для быстрого поиска
    b_cleaned_to_original = defaultdict(list)
    for idx, row in df_data_b.iterrows():
        b_cleaned_to_original[row['cleaned']].append(row['data2'])

    b_cleaned_list = list(b_cleaned_to_original.keys())
    
    # Предварительная обработка для rapidfuzz
    processed_b = [utils.default_process(x) for x in b_cleaned_list]
    
    results = []
    
    # Загружаем конфигурацию для выбора метрики
    config = load_config()
    use_token_sort = config.get("comparison_options", {}).get("use_token_sort_ratio", 0) == 1
    
    # Выбираем скорер
    scorer = fuzz.token_sort_ratio if use_token_sort else fuzz.ratio
    
    for idx, row in enumerate(df_data_a.itertuples(), 1):
        cleaned_a = row.cleaned
        processed_a = utils.default_process(cleaned_a)
        
        # Используем rapidfuzz с выбранным скорером
        matches = process.extract(
            processed_a, 
            processed_b, 
            scorer=scorer, # Используем выбранный скорер
            score_cutoff=similarity_criterion,
            limit=50  # Ограничиваем количество результатов
        )
        
        matched_strings = []
        for match_text, score, match_idx in matches:
            if score >= similarity_criterion:
                original_strings = b_cleaned_to_original[b_cleaned_list[match_idx]]
                matched_strings.extend(original_strings)
        
        if matched_strings:
            results.append({
                'data1': row.data1,
                'data2': matched_strings
            })
        
    # Создаем финальный DataFrame одним действием
    if results:
        df_output = pd.DataFrame([
            {'data1': result['data1'], 'data2': data2} 
            for result in results 
            for data2 in result['data2']
        ]).drop_duplicates()
        
        file_path = os.path.join(script_dir, "working_files", NAME_OUTPUT_FILE)
        try:
            df_output.to_excel(file_path, index=False)
        except ValueError:
            # pr_bar.update(progress=0, total=None)
            # label_progress_bar.update(content = '')
            raise Sheet_too_large_Error()
        except PermissionError:
            # pr_bar.update(progress=0, total=None)
            # label_progress_bar.update(content = '')
            raise PermissionError()
            
    else:
        # Создаем пустой файл если нет результатов
        file_path = os.path.join(script_dir, "working_files", NAME_OUTPUT_FILE)
        pd.DataFrame(columns=['data1', 'data2']).to_excel(file_path, index=False)

    # Очищаем кэш после использования
    _cleaning_cache.clear()