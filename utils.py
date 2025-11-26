import json
import re
from typing import Dict, Any
from nltk.corpus import stopwords

# Глобальная переменная для кэширования конфигурации
_config_cache: Dict[str, Any] = {}
CONFIG_FILE_PATH = "config.json"

RUSSIAN_STOPWORDS = set(stopwords.words('russian'))

# Значения по умолчанию для config.json
# Используем значения из config.json, который вы предоставили в первом сообщении
DEFAULT_CONFIG = {
    "cleaning_options": {
        "use_stemming_or_lemmatization": 1,
        "remove_legal_forms": 0,
        "remove_digits": 0, 
        "remove_punctuation": 1,
        "remove_stopwords": 1,
        "sort_words": 1
    },
    "comparison_options": {
        "use_token_sort_ratio": 1,
        "similarity_score": 90
    },
    "legal_forms_regex": "(?i)\\b(ООО|ОАО|АО|ЗАО|ПФ|ПАО|L.L.C|ИП|ТОО|Ltd|Co.НП|СО|КП|ФК|ГК|ЗАО|ОАО|ПАО|ИП|ТОО|LLP|PLC|S.A.|S.R.L.|GmbH|B.V.|Inc.|Corp.|S.p.A.|Pty Ltd|SAS|N.V.)\\b"
}

def write_default_config(config_path: str = None):
    """Создает файл config.json со значениями по умолчанию."""
    if config_path is None:
        config_path = CONFIG_FILE_PATH
    
    try:
        with open(config_path, 'w', encoding='utf-8') as file:
            json.dump(DEFAULT_CONFIG, file, ensure_ascii=False, indent=4)
        # print(f"Создан файл конфигурации по умолчанию: {config_path}")
    except Exception as e:
        print(f"Ошибка при создании файла конфигурации по умолчанию: {e}")

def read_config(config_path: str = None) -> dict:
    """
    Читает и возвращает текущую конфигурацию из JSON файла.
    Если файл не найден или некорректен, создает его со значениями по умолчанию.
    """
    if config_path is None:
        config_path = CONFIG_FILE_PATH
    
    # 1. Попытка прочитать файл
    try:
        with open(config_path, 'r', encoding='utf-8') as file:
            config = json.load(file)
            return config
    
    # 2. Обработка FileNotFoundError: файл не найден
    except FileNotFoundError:
        # print(f"Файл конфигурации {config_path} не найден. Создание файла по умолчанию.")
        write_default_config(config_path)
        return DEFAULT_CONFIG
        
    # 3. Обработка json.JSONDecodeError: некорректный формат
    except json.JSONDecodeError:
        # print(f"Ошибка: Неверный формат JSON в файле {config_path}. Файл будет перезаписан значениями по умолчанию.")
        write_default_config(config_path)
        return DEFAULT_CONFIG
        
    # 4. Обработка других ошибок
    except Exception as e:
        print(f"Непредвиденная ошибка при чтении конфигурации: {e}. Возврат значений по умолчанию.")
        return DEFAULT_CONFIG 

def update_config(updates, config_path: str = None):
    
    """
    Вносит изменения в файл конфигурации JSON
    
    Args:
        config_path (str): Путь к файлу config.json
        updates (dict): Словарь с обновлениями для конфигурации
    """
    
    if config_path is None:
        config_path = CONFIG_FILE_PATH
    
    # Сначала читаем конфигурацию, которая теперь гарантированно вернет либо существующую, либо дефолтную
    config = read_config(config_path)
    
    try:
        # Рекурсивное обновление конфигурации
        def deep_update(current_dict, update_dict):
            for key, value in update_dict.items():
                if (key in current_dict and 
                    isinstance(current_dict[key], dict) and 
                    isinstance(value, dict)):
                    deep_update(current_dict[key], value)
                else:
                    current_dict[key] = value
        
        # Применение обновлений
        deep_update(config, updates)
        
        # Запись обновленной конфигурации
        with open(config_path, 'w', encoding='utf-8') as file:
            json.dump(config, file, ensure_ascii=False, indent=4)
        
        # print("Конфигурация успешно обновлена")
        clear_config_cache()
        return True
        
    except Exception as e:
        # print(f"Ошибка при обновлении конфигурации: {e}")
        return False

def load_config(file_path: str = CONFIG_FILE_PATH) -> Dict[str, Any]:
    global _config_cache
    if not _config_cache:
        _config_cache = read_config(file_path)
    return _config_cache

# Инициализация лемматизатора (если pymorphy2 не установлен, будет использоваться заглушка)
try:
    from pymorphy3 import MorphAnalyzer
    morph = MorphAnalyzer()
    def lemmatize_word(word: str) -> str:
        return morph.parse(word)[0].normal_form
except ImportError:
    print("Предупреждение: pymorphy2 не установлен. Лемматизация будет пропущена.")
    def lemmatize_word(word: str) -> str:
        return word # Заглушка

def clean_text_optimized(text: str, config: Dict[str, Any]) -> str:
    """
    Оптимизированная функция очистки текста с использованием опциональных параметров.
    """
    options = config.get("cleaning_options", {})
    text = text.lower()
    
    # 1. Удаление организационно-правовых форм (ОПФ)
    if options.get("remove_legal_forms", 0) == 1:
        regex = config.get("legal_forms_regex", "")
        if regex:
            text = re.sub(regex, ' ', text)
            
    # 2. Удаление цифр
    if options.get("remove_digits", 0) == 1:
        text = re.sub(r'\d+', ' ', text)


        
    # 3. Удаление знаков препинания (кроме тех, что могут быть частью слова, если не удалять все)
    # В текущем коде cleantext.clean(punct=True) удаляет все. Повторяем это поведение.
    if options.get("remove_punctuation", 1) == 1:
        # Удаляем все, кроме букв и пробелов
        text = re.sub(r'[^\w\s]', ' ', text)
        
    # 4. Удаление лишних пробелов (аналог extra_spaces=True)
    text = re.sub(r'\s+', ' ', text).strip()
    
    # 5. Токенизация и обработка слов
    words = text.split()
    
    # 6. Удаление стоп-слов
    if options.get("remove_stopwords", 1) == 1:
        words = [word for word in words if word not in RUSSIAN_STOPWORDS]
        
    # 7. Лемматизация/Стемминг
    if options.get("use_stemming_or_lemmatization", 0) == 1:
        words = [lemmatize_word(word) for word in words]
        
    # 8. Сортировка слов
    if options.get("sort_words", 1) == 1:
        words.sort()
        
    return " ".join(words)

# Добавляем функцию для очистки кэша, чтобы можно было перечитать конфиг
def clear_config_cache():
    global _config_cache
    _config_cache = {}
