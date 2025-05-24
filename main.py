# -*- coding: utf-8 -*-
"""
Created on Fri May 16 09:59:06 2025

@author: a.karabedyan
"""

import pandas as pd
import os

from textual import work
from textual.app import App, ComposeResult
from textual.containers import Horizontal
from textual.widgets import Header, Footer, Static, LoadingIndicator, ProgressBar, MaskedInput
from textual.containers import Middle, Center

from comparison import create_file_matches
from data_text import (correct_columns,
                       NAME_APP,
                       NAME_DATA_FILE,
                       NAME_OUTPUT_FILE,
                       SUB_TITLE_APP,
                       EXAMPLE,
                       TEXT_INTRODUCTION,
                       TEXT_CREATE_DATACOMPARISON_FILE,
                       TEXT_CREATE_FUZZYMAPPINGRESULT_FILE,
                       TEXT_WAITING_FUZZYMAPPINGRESULT_FILE,
                       TEXT_MISSING_COL,
                       TEXT_MISSING_DATA_COL,
                       TEXT_END,
                       TEXT_ERR_CREATE_FUZZYMAPPINGRESULT_FILE,
                       TEXT_ERR_VALID_SIMILARITI_SCORE,
                       TEXT_DATACOMPARISON_FILE_NOT_FIND,
                       TEXT_ERR_PERMISSION,
                       TEXT_APP_EXCEL_NOT_FIND,
                       TEXT_UNKNOW_ERR)

class FuzzyMatchToolApp(App):
    CSS = """
    .introduction {
        height: auto;
        border: solid #0087d7;
    }
    .steps_l {
        height: auto;
        width: 9fr;
    }
    .steps_r {
        height: 100%;
        border: solid #0087d7;
        width: 1fr;
    }
    Horizontal {
        height: auto;
        border: solid #0087d7;
    }
    LoadingIndicator {
        dock: bottom;
        height: 10%;
    }
    ProgressBar {
        padding-left: 3;
    }
    """

    BINDINGS = [
        ("ctrl+r", "open_file", "Открыть файл с исходными данными для редактирования"),
        ("ctrl+m", "open_comparison", "Сформировать и открыть файл с совпадениями"),
    ]

    def compose(self) -> ComposeResult:
        yield Header(show_clock=True, icon='<>')
        yield Static(TEXT_INTRODUCTION, classes='introduction')
        yield Horizontal(
                Static(TEXT_CREATE_DATACOMPARISON_FILE, classes='steps_l'),
                MaskedInput(template="99", value="90", classes='steps_r'), id='example')
        yield Footer()
        yield LoadingIndicator()
        with Middle():
            with Center():
                yield ProgressBar(show_eta=False)

    def on_mount(self) -> None:
        self.title = NAME_APP
        self.sub_title = SUB_TITLE_APP
        self.query_one(ProgressBar).visible = False
        self.query_one(LoadingIndicator).visible = False

    '''
    Открывает excel-файл, в два столбца которого пользователь вставляет текстовые
    данные для сравнения между собой. Файл будет содержать образец для заполнения.
    Файл формируется из датафрейма pandas, который в свою очередь формируется из
    словаря EXAMPLE.
    '''
    @work(thread=True)
    def action_open_file(self) -> None:
        self.query_one(LoadingIndicator).visible = True
        self.query_one('.steps_l').update(TEXT_CREATE_FUZZYMAPPINGRESULT_FILE)

        if not os.path.isfile(NAME_DATA_FILE):
            example_file = pd.DataFrame(EXAMPLE)
            example_file.to_excel(NAME_DATA_FILE, index=False)
        try:
            os.startfile(NAME_DATA_FILE)
        except OSError as e:
            error_message = TEXT_APP_EXCEL_NOT_FIND.format(error_app_xls=e,
                                                           NAME_DATA_FILE=NAME_DATA_FILE)
            self.query_one('.steps_l').update(error_message)
        self.query_one(LoadingIndicator).visible = False


    '''
    Формирует с помощью функции create_file_matches и открывает excel-файл,
    содержащий похожие строки, соответствующие уровню схожести, установленному
    пользователем (90% по умолчанию).
    '''

    @work(thread=True)
    def action_open_comparison(self) -> None:
        self.query_one(LoadingIndicator).visible = True
        self.query_one(MaskedInput).visible = False
        try:
            df = pd.read_excel(NAME_DATA_FILE)

            # Проверка наличия необходимых столбцов
            if not set(correct_columns).issubset(set(df.columns)):
                missing_columns = set(correct_columns) - set(df.columns)
                missing_columns_str = ', '.join(missing_columns)
                error_message = TEXT_MISSING_COL.format(name_data_file=NAME_DATA_FILE,
                                                        missing_columns_str=missing_columns_str)
                self.query_one('.steps_l').update(error_message)
            # Проверка заполненности первой строки
            elif df.iloc[0].isnull().any():
                null_columns = df.iloc[0].isnull()
                missing_columns = null_columns[null_columns].index.tolist()
                missing_columns_str = ', '.join(missing_columns)
                error_message = TEXT_MISSING_DATA_COL.format(name_data_file=NAME_DATA_FILE,
                                                             missing_columns_str=missing_columns_str)
                self.query_one('.steps_l').update(error_message)
            else:
                try:
                    similarity_score = int(self.query_one(MaskedInput).value)
                    if 10 <= similarity_score <= 99:
                        self.query_one(ProgressBar).update(total=100)
                        self.query_one(ProgressBar).visible = True
                        self.query_one('.steps_l').update(TEXT_WAITING_FUZZYMAPPINGRESULT_FILE)

                        create_file_matches(self.query_one(ProgressBar), similarity_score)

                        self.query_one(ProgressBar).visible = False
                        self.query_one(ProgressBar).update(progress=0)

                        if os.path.isfile(NAME_OUTPUT_FILE):
                            os.startfile(os.path.abspath(NAME_OUTPUT_FILE))  # Убедитесь, что Вы на Windows
                            self.query_one('.steps_l').update(TEXT_END)
                        else:
                            self.query_one('.steps_l').update(TEXT_ERR_CREATE_FUZZYMAPPINGRESULT_FILE)
                    else:
                        self.query_one('.steps_l').update(TEXT_ERR_VALID_SIMILARITI_SCORE)
                except ValueError:
                    self.query_one('.steps_l').update(TEXT_ERR_VALID_SIMILARITI_SCORE)
                except PermissionError:
                    self.query_one('.steps_l').update(TEXT_ERR_PERMISSION)
                    self.query_one(ProgressBar).visible = False
                    self.query_one(ProgressBar).update(progress=0)
                except Exception as e:
                    self.query_one('.steps_l').update(TEXT_UNKNOW_ERR.format(text_err=e))
        except FileNotFoundError:
            self.query_one('.steps_l').update(TEXT_DATACOMPARISON_FILE_NOT_FIND)
        except ValueError as e:
        # except Exception as e:
            self.query_one('.steps_l').update(TEXT_UNKNOW_ERR.format(text_err=e))
        finally:
            self.query_one(LoadingIndicator).visible = False
            self.query_one(MaskedInput).visible = True
            self.query_one(ProgressBar).visible = False

if __name__ == "__main__":
    app = FuzzyMatchToolApp()
    app.run()
