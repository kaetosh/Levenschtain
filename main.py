import os
import pandas as pd
from textual import work
from textual.app import App, ComposeResult
from textual.widgets import Button, Header, Footer, Markdown, MaskedInput, Static, LoadingIndicator, Switch
from textual.containers import Horizontal, Container
from textual.screen import ModalScreen
from textual.binding import Binding
from textual.validation import Number
from textual.reactive import reactive

# Assuming these modules are available and contain the necessary constants/functions
# In a real-world scenario, I would also refactor these modules.
from text import TEXT_BRIEF_INTRODUCTION, TEXT_HELP, EXAMPLE, NAME_DATA_FILE, correct_columns, NAME_OUTPUT_FILE
from comparison import create_file_matches
from custom_errors import Sheet_too_large_Error
from utils import update_config, read_config, DEFAULT_CONFIG

# --- Utility Functions ---

def check_sheet_exists(file_path: str, sheet_name: str = 'Sheet1') -> bool:
    """
    Checks if a specific sheet exists in an Excel file without reading the whole file.
    
    Args:
        file_path: Path to the Excel file.
        sheet_name: Name of the sheet to check.
        
    Returns:
        True if the sheet exists, False otherwise.
    """
    try:
        # Use read_excel with sheet_name=None to get a dictionary of all sheets
        # and check if the target sheet is in the keys. This is generally faster
        # than trying to read the sheet and catching an exception.
        excel_file = pd.ExcelFile(file_path)
        return sheet_name in excel_file.sheet_names
    except ValueError as e:
        # Catch pandas ValueError if the file is not a valid Excel file
        # or if there's a more general read error.
        if "Worksheet" in str(e) and "not found" in str(e):
            return False
        # Re-raise other ValueErrors
        raise
    except FileNotFoundError:
        return False
    except Exception:
        # Catch other potential errors during file reading (e.g., corrupted file)
        return False

# --- Modal Screens ---

class SettingsScreen(ModalScreen):
    """
    Окно с настройками для параметров очистки и нормализации данных.
    """
    
    def compose(self) -> ComposeResult:
        # 1. Читаем конфигурацию при создании окна
        config = read_config()
        
        # Устанавливаем значения по умолчанию
        cleaning_options = config.get("cleaning_options", DEFAULT_CONFIG.get("cleaning_options", {}))
        
        # 2. Извлекаем значения и преобразуем их в булевы для Switch
        # Логика инверсии:
        # - use_stemming_or_lemmatization: 1 (вкл) -> True (Switch вкл)
        # - remove_*: 1 (удалить) -> False (Switch выкл, т.е. "не учитывать")
        
        lemming_value = bool(cleaning_options.get("use_stemming_or_lemmatization", 1))
        digits_value = not bool(cleaning_options.get("remove_digits", 0))
        forms_value = not bool(cleaning_options.get("remove_legal_forms", 1))
        sort_value = not bool(cleaning_options.get("sort_words", 1))
        
        yield Container(
            Horizontal(
                Static("Нормализация слов:", classes="statics-settings-modal"),
                Switch(value=lemming_value, id='switch-lemming', classes="switchs-settings-modal"),
                classes="horizontals-settings-modal",
                id='horizontal-lemming-settings-modal'
                ),
            Horizontal(
                Static("Учитывать порядок слов:", classes="statics-settings-modal"),
                Switch(value=sort_value, id='switch-sort', classes="switchs-settings-modal"),
                classes="horizontals-settings-modal",
                id='horizontal-sort-settings-modal'
                ),
            Horizontal(
                Static("Учитывать цифры:", classes="statics-settings-modal"),
                Switch(value=digits_value, id='switch-digits', classes="switchs-settings-modal"),
                classes="horizontals-settings-modal",
                id='horizontal-digits-settings-modal'
                ),
            Horizontal(
                Static("Учитывать аббревиатуры формы юр.лиц:", classes="statics-settings-modal"),
                Switch(value=forms_value, id='switch-forms', classes="switchs-settings-modal"),
                classes="horizontals-settings-modal",
                id='horizontal-forms-settings-modal'
                ),
            Horizontal(
                Button("Сохранить", variant="success", id="button-settings-modal"),
                id="horizontals-button-settings-modal"),
            id="container-settings-modal"
        )
    
    def on_mount(self) -> None:
        self.query_one('#horizontal-lemming-settings-modal').tooltip = 'Приведение слов исходного текста к словарной форме'
        self.query_one('#horizontal-sort-settings-modal').tooltip = 'Важна ли последовательность слов в исходном тексте'
        self.query_one('#horizontal-digits-settings-modal').tooltip = 'Учитывать числовые значения в исходном тексте'
        self.query_one('#horizontal-forms-settings-modal').tooltip = 'различать ООО/ЗАО/ИП и т.д.'
    
    def on_button_pressed(self, event: Button.Pressed):
        """Обрабатывает нажатие кнопки "Сохранить"."""
        if event.button.id == "button-settings-modal":
            
            # Собираем значения из виджетов Switch
            # True (Switch вкл) -> 1 (в config), False (Switch выкл) -> 0 (в config)
            lemming_val = int(self.query_one('#switch-lemming', Switch).value)
            
            # Инверсия: True (учитывать/Switch вкл) -> 0 (не удалять/в config), 
            # False (не учитывать/Switch выкл) -> 1 (удалить/в config)
            remove_digits_val = int(not self.query_one('#switch-digits', Switch).value)
            remove_forms_val = int(not self.query_one('#switch-forms', Switch).value)
            sort_words_val = int(not self.query_one('#switch-sort', Switch).value)
            
            updates = {
                "cleaning_options": {
                    "use_stemming_or_lemmatization": lemming_val,
                    "remove_digits": remove_digits_val,
                    "remove_legal_forms": remove_forms_val,
                    "sort_words": sort_words_val,
                },
                # 'use_token_sort_ratio' зависит от 'sort_words'
                "comparison_options": {
                    "use_token_sort_ratio": sort_words_val,
                }
            }
            
            # Обновляем конфигурацию
            update_config(updates=updates)
            
            # Закрываем модальное окно
            self.app.pop_screen()


class HelpScreen(ModalScreen):
    """Экран со справкой"""
    
    BINDINGS = [("escape", "dismiss", "Закрыть")] # Используем dismiss вместо action_close
    
    def compose(self) -> ComposeResult:
        yield Markdown(TEXT_HELP)
        yield Footer()
    
    def on_mount(self) -> None:
        self.border_title = "Справка - нажмите Escape для закрытия"

class SetSimilarityLevelScreen(ModalScreen):
    """Экран установки параметра схожести пользователем"""
    
    
    def compose(self) -> ComposeResult:
        # 1. Читаем конфигурацию при создании окна
        config = read_config()
        
        # Устанавливаем значения по умолчанию
        comparison_options = config.get("comparison_options", {})
        
        # Используем int() для безопасности, затем str() для MaskedInput
        similarity_score = str(int(comparison_options.get("similarity_score", 90)))
        input_validator = Number(minimum=10, maximum=99)
        
        yield Container(
            Static("Введите степень схожести в процентах от 10 до 99",
                   id="static-set-similarity-level-modalscreen"),
            Horizontal(
                MaskedInput(template='99',
                            value=similarity_score,
                            validators=[input_validator],
                            id="input-set-similarity-level-modalscreen"),
                Button("Сохранить", variant="success", id="button-set-similarity-level-modalscreen"),
                id="horizontal-set-similarity-level-modalscreen"),
            id="container-set-similarity-level-modalscreen"
            )
    
    def on_input_changed(self, event: MaskedInput.Changed) -> None:
        """Обрабатывает изменение ввода для включения/выключения кнопки 'Сохранить'."""
        save_button = self.query_one('#button-set-similarity-level-modalscreen', Button)
        # Проверяем, что validation_result существует и не является валидным
        is_valid = event.validation_result and event.validation_result.is_valid
        save_button.disabled = not is_valid
    
    def on_button_pressed(self, event: Button.Pressed):
        if event.button.id == "button-set-similarity-level-modalscreen":
            input_widget = self.query_one(MaskedInput)
            new_score = input_widget.value
            
            updates = {
                "comparison_options": {
                    "similarity_score": new_score,
                }
            }
            # Обновляем конфигурацию
            update_config(updates=updates)
            
            # Обновляем reactive-переменную в главном приложении
            app = self.app
            app.similarity_score = new_score
            
            # Обновляем tooltip на главной странице
            app.query_one("#set_similarity_level").tooltip = f"Текущее значение параметра {new_score}%"
            self.app.pop_screen()


# --- Main Application ---

class FuzzyMatchToolApp(App[str]):
    # CSS_PATH = "stile.tcss"
    TITLE = "FuzzyMatchTool"
    SUB_TITLE = "мастер по поиску совпадений"
    
    
    CSS = """
    Header {
      dock: top;
      content-align: center middle;
    }

    #introduction {
        border: tall $background;
        content-align: center top;
        border: solid $accent;
        width: 100%;
        height: 19;
    }

    #horizontal_progress_bar {
        align: center bottom;
        width: 100%;
        height: 1;
    }

    #buttons {
        align: center middle;
        layout: grid;
        grid-size: 3 1;
        height: 5;
    }
    Button {
        border: tall $background;
        width: 100%;
    }



    SetSimilarityLevelScreen {
            align: center middle;
        }
            #container-set-similarity-level-modalscreen {
               align: center middle;
               width: 70; 
               height: 10;
               border: solid $accent;
               background: $surface;
               padding: 1;
               margin: 1;
            }
            
            #static-set-similarity-level-modalscreen {
                content-align: center middle;
            }
            
            #horizontal-set-similarity-level-modalscreen {
               align: center middle; 
            }
            
            #input-set-similarity-level-modalscreen {
                width: 15%;
                content-align: center middle;
            }


            #button-set-similarity-level-modalscreen {
                width: 25%;
                }
                
    SettingsScreen {
            align: center middle;
        }
            #container-settings-modal {
               width: 55; 
               height: 20;
               border: solid $accent;
               background: $surface;
               padding: 1;
            }
           .statics-settings-modal {
              content-align: left middle;
              height: 3;
              width: 80%; 
           }
           .switchs-settings-modal {
               content-align: right middle;
               width: 20%;
           }
           #horizontals-button-settings-modal{
               align: center bottom;
               }
           #button-settings-modal {
                width: 45%;
                        }
    """
    
    
    # Используем action_dismiss для закрытия модальных окон
    BINDINGS = [
        Binding(key="f1", action="push_screen('help')", description="Помощь", key_display="F1"),
        Binding(key="f2", action="open_dir", description="Открыть папку с файлами", key_display="F2"),
        Binding(key="f3", action="push_screen('settings')", description="Настройки", key_display="F3")
    ]
    # Инициализация reactive-переменной из конфига
    config = read_config()
    similarity_score = reactive(str(config.get("comparison_options", {}).get("similarity_score", '99')))
    
    def compose(self) -> ComposeResult:
        markdown = Markdown(TEXT_BRIEF_INTRODUCTION, id="introduction")
        markdown.code_indent_guides = False
        
        # Убираем show_command_palette = False, так как это не нужно для Footer
        yield Header()
        yield markdown
        yield Horizontal(
            Button("📋 Открыть исходные данные", id="open_data", variant="primary"),
            Button("🔧 Установить уровень схожести", id="set_similarity_level", variant="primary"),
            Button("🔍 Найти схожие значения", id="find_similar_values", variant="primary"),
            id="buttons")
        yield Horizontal(LoadingIndicator(), id='horizontal_progress_bar')
        yield Footer(show_command_palette = False)

    def on_mount(self) -> None:
        self.query_one(LoadingIndicator).visible = False
        # Регистрация экранов
        self.install_screen(HelpScreen(), name="help")
        self.install_screen(SettingsScreen(), name="settings")
        self.install_screen(SetSimilarityLevelScreen(), name="similarity_level")
        
        # Установка начального tooltip (используем .value для получения строки)
        
        self.query_one("#set_similarity_level", Button).tooltip = f"Текущее значение параметра {self.similarity_score}%"

    def on_button_pressed(self, event: Button.Pressed) -> None:
        
        if event.button.id == "open_data":
            self.create_and_open_excel_template()
            
        elif event.button.id == "set_similarity_level":
            self.push_screen("similarity_level")
            
        elif event.button.id == "find_similar_values":
            self.start_comparison_process()
            
    # Упрощенные action-методы для использования push_screen с именем
    def action_push_screen(self, screen_name: str) -> None:
        """Действие для открытия экрана по имени."""
        self.push_screen(screen_name)

    def action_open_dir(self) -> None:
        """Действие при нажатии F2 - открывает рабочую папку."""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        working_dir = os.path.join(script_dir, "working_files")
        
        os.makedirs(working_dir, exist_ok=True)
        
        # Используем os.startfile только для Windows. 
        # Для кросс-платформенности лучше использовать 'xdg-open' (Linux) или 'open' (macOS)
        # Но для простоты и соответствия исходному коду оставляем os.startfile
        try:
            os.startfile(working_dir)
        except AttributeError:
            # Fallback for non-Windows systems (e.g., using 'open' or 'xdg-open')
            import subprocess
            if os.name == 'posix': # Linux, macOS
                subprocess.run(['xdg-open', working_dir] if 'linux' in os.uname().sysname.lower() else ['open', working_dir])
            else:
                self.notify("Не удалось открыть папку автоматически. Рабочая папка: working_files", severity='warning')

    def finish_processing(self):
        """Сбрасывает состояние интерфейса после завершения обработки."""
        self.query_one(LoadingIndicator).visible = False
        # Используем query_many для повышения производительности
        for button in self.query("Button"):
            button.disabled = False
    
    def create_and_open_excel_template(self):
        """
        Создает (если не существует) и открывает файл-шаблон для ввода данных.
        """
        script_dir = os.path.dirname(os.path.abspath(__file__))
        working_dir = os.path.join(script_dir, "working_files")
        os.makedirs(working_dir, exist_ok=True)
        file_path = os.path.join(working_dir, NAME_DATA_FILE)
        
        if not os.path.isfile(file_path):
            example_file = pd.DataFrame(EXAMPLE)
            example_file.to_excel(file_path, index=False)
                
        self.notify('Открываем файл.', title="Информация", severity='information', timeout=2)
        try:
            os.startfile(file_path)
        except OSError:
            self.notify('Убедитесь, что Excel установлен и попробуйте снова.',
                            title="Ошибка",
                            severity='error',
                            timeout=5)
            
    def start_comparison_process(self):
        """Инициализирует процесс сравнения, блокирует UI и запускает worker."""
        self.query_one(LoadingIndicator).visible = True
        for button in self.query("Button"):
            button.disabled = True
        
        # Запускаем worker
        self.create_and_open_excel_comparison()

    @work(thread=True)
    def create_and_open_excel_comparison(self):
        """
        Worker-метод: выполняет сравнение данных и открывает файл с результатами.
        """
        
        script_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(script_dir, "working_files", NAME_DATA_FILE)

        if not os.path.isfile(file_path):
            self.notify(f'Отсутствует {NAME_DATA_FILE}. Нажмите 📋 Открыть исходные данные, чтобы сформировать и заполнить файл',
                            title="Ошибка",
                            severity='error',
                            timeout=10)
            self.call_later(self.finish_processing)
            return
            
        if not check_sheet_exists(file_path):
            self.notify(f'В {NAME_DATA_FILE} отсутствует лист Sheet1.',
                            title="Ошибка",
                            severity='error',
                            timeout=5)
            self.call_later(self.finish_processing)
            return
            
        try:
            # Чтение данных
            df = pd.read_excel(io=file_path, sheet_name='Sheet1')
        except Exception as e:
            self.notify(f"Ошибка чтения файла: {e}", title="Ошибка", severity='error', timeout=10)
            self.call_later(self.finish_processing)
            return

        # --- Валидация данных ---
        
        # 1. Проверка количества столбцов
        if len(df.columns) != 2:
            self.notify(f"""\
В {NAME_DATA_FILE} должно быть только два столбца.
Нажмите 📋 Открыть исходные данные, чтобы открыть для редактирования {NAME_DATA_FILE}""",
                            title="Ошибка",
                            severity='error',
                            timeout=10)
            self.call_later(self.finish_processing)
            return
            
        # 2. Проверка наличия необходимых столбцов
        if not set(correct_columns).issubset(set(df.columns)):
            missing_columns = set(correct_columns) - set(df.columns)
            missing_columns_str = ', '.join(missing_columns)
            error_message = f"""\
В {NAME_DATA_FILE} ошибки в именах столбцов: {missing_columns_str}.
Нажмите 📋 Открыть исходные данные, чтобы открыть для редактирования {NAME_DATA_FILE}"""
            self.notify(error_message,
                            title="Ошибка",
                            severity='error',
                            timeout=10)
            self.call_later(self.finish_processing)
            return
        # 3. Проверка заполненности первой строки (проверка на пустой файл)
        if df.empty or df.iloc[0].isnull().all():
            self.notify(f"Файл {NAME_DATA_FILE} возможно пуст или не содержит данных. Проверьте, заполнена ли первая строка таблицы.",
                            title="Ошибка",
                            severity='error',
                            timeout=10)
            self.call_later(self.finish_processing)
            return
        
        # 4. Проверка на пустые значения в первой строке (как в оригинале)
        if df.iloc[0].isnull().any():
            null_columns = df.iloc[0].isnull()
            missing_columns = null_columns[null_columns].index.tolist()
            missing_columns_str = ', '.join(missing_columns)
            error_message = f"""\
В {NAME_DATA_FILE} не заполнена первая строка таблицы в следующих столбцах: {missing_columns_str}.
Нажмите 📋 Открыть исходные данные, чтобы открыть для редактирования {NAME_DATA_FILE}"""
            self.notify(error_message,
                            title="Ошибка",
                            severity='error',
                            timeout=10)
            self.call_later(self.finish_processing)
            return
        
        
        
        # --- Обработка данных ---
        
        self.notify('Начинаем обработку данных.',
                        title="Информация",
                        severity='information',
                        timeout=2)
        try:
            # Используем int(self.similarity_score) для получения актуального значения
            create_file_matches(int(self.similarity_score))
            
            output_file_path = os.path.join(script_dir, "working_files", NAME_OUTPUT_FILE)
            
            if os.path.isfile(output_file_path):
                self.notify('Открываем файл с результатами сравнения.',
                                title="Информация",
                                severity='information',
                                timeout=2)
                try:
                    os.startfile(os.path.abspath(output_file_path))
                except PermissionError:
                    self.notify(f"Нет доступа к {NAME_OUTPUT_FILE}. Пожалуйста, закройте данный файл и нажмите 🔍 Найти схожие значения.",
                                    title="Ошибка",
                                    severity='error',
                                    timeout=5)
                except OSError:
                    self.notify(f"Не удалось открыть файл {NAME_OUTPUT_FILE} автоматически.",
                                    title="Ошибка",
                                    severity='warning',
                                    timeout=5)
            else:
                self.notify(f"Ошибка при формировании {NAME_OUTPUT_FILE}",
                                title="Ошибка",
                                severity='error',
                                timeout=5)
        
        except Sheet_too_large_Error:
            self.notify("Найденных совпадений больше строк в excel.",
                            title="Ошибка",
                            severity='error',
                            timeout=5)
        except PermissionError:
            self.notify(f"Нет доступа к {NAME_OUTPUT_FILE}. Пожалуйста, закройте данный файл и нажмите 🔍 Найти схожие значения.",
                            title="Ошибка",
                            severity='error',
                            timeout=5)
        except Exception as e:
            self.notify(f"Непредвиденная ошибка во время обработки: {e}",
                            title="Критическая ошибка",
                            severity='error',
                            timeout=10)
            
        self.call_later(self.finish_processing)


if __name__ == "__main__":
    # Убеждаемся, что app создается только один раз
    app = FuzzyMatchToolApp()
    app.run()
