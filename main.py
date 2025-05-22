# -*- coding: utf-8 -*-
"""
Created on Fri May 16 09:59:06 2025

@author: a.karabedyan
"""

import pandas as pd
import os
import asyncio

from fuzzywuzzy import process, fuzz
from cleantext import clean
import re
from typing import List, Tuple

from textual.app import App, ComposeResult
from textual.widgets import Header, Footer, Static, LoadingIndicator, ProgressBar
from textual.reactive import reactive
from textual import work
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
                       TEXT_DATACOMPARISON_FILE_NOT_FIND,
                       TEXT_APP_EXCEL_NOT_FIND)

class HeaderApp(App):
    CSS = """
    .introduction {
        height: auto;
        border: solid green;
    }
    .steps {
        height: auto;
        border: solid green;
    }
    .indicator {
        dock: bottom;
        height: auto;
    }
    ProgressBar {
        padding-left: 3;

    }
    """

    BINDINGS = [
        ("ctrl+r", "open_file", "Открыть data_comparison.xlsx для редактирования"),
        ("ctrl+d", "open_comparison", "Сформировать и открыть файл с совпадениями"),
    ]

    loading_indicator = reactive(None)

    def compose(self) -> ComposeResult:
        self.loading_indicator = LoadingIndicator(classes='indicator')
        yield Header(show_clock=True, icon='<>')
        yield Footer()
        yield Static(TEXT_INTRODUCTION, classes='introduction')
        yield Static(TEXT_CREATE_DATACOMPARISON_FILE, id='example', classes='steps')
        yield self.loading_indicator
        with Middle():
            with Center():
                yield ProgressBar(show_eta=False)

    def on_mount(self) -> None:
        self.title = NAME_APP
        self.sub_title = SUB_TITLE_APP
        self.query_one(ProgressBar).visible = False
        self.loading_indicator.visible = False


    @work(thread=True)
    def action_open_file(self) -> None:
        self.loading_indicator.visible = True
        central_widget = self.query_one('#example')
        central_widget.update(TEXT_CREATE_FUZZYMAPPINGRESULT_FILE)

        if not os.path.isfile(NAME_DATA_FILE):
            example_file = pd.DataFrame(EXAMPLE)
            example_file.to_excel(NAME_DATA_FILE, index=False)
        try:
            os.startfile(NAME_DATA_FILE)
        except OSError as e:
            error_message = TEXT_APP_EXCEL_NOT_FIND.format(error_app_xls=e,
                                                           NAME_DATA_FILE=NAME_DATA_FILE)
            central_widget.update(error_message)
        self.loading_indicator.visible = False

    @work(thread=True)
    def action_open_comparison(self) -> None:
        self.loading_indicator.visible = True

        central_widget = self.query_one('#example')

        try:
            df = pd.read_excel(NAME_DATA_FILE)
            if set(correct_columns) != set(df.columns):
                missing_columns = set(correct_columns) - set(df.columns)
                missing_columns_str = ', '.join(missing_columns)
                error_message = TEXT_MISSING_COL.format(name_data_file=NAME_DATA_FILE,
                                                        missing_columns_str=missing_columns_str)
                self.query_one('#example').update(error_message)
            elif df.iloc[0].isnull().any():
                null_columns = df.iloc[0].isnull()
                missing_columns = null_columns[null_columns].index.tolist()
                missing_columns_str = ', '.join(missing_columns)
                error_message = TEXT_MISSING_DATA_COL.format(name_data_file=NAME_DATA_FILE,
                                                             missing_columns_str=missing_columns_str)
                self.query_one('#example').update(error_message)
            else:
                self.query_one(ProgressBar).update(total=100)
                self.query_one(ProgressBar).visible = True
                self.query_one('#example').update(TEXT_WAITING_FUZZYMAPPINGRESULT_FILE)

                create_file_matches(self.query_one(ProgressBar))

                self.query_one(ProgressBar).visible = False
                self.query_one(ProgressBar).update(progress=0)
                if os.path.isfile(NAME_OUTPUT_FILE):
                    os.startfile(f'file://{os.path.abspath(NAME_OUTPUT_FILE)}')
                    central_widget.update(TEXT_END)
                else:
                    self.query_one('#example').update(TEXT_ERR_CREATE_FUZZYMAPPINGRESULT_FILE)
        except FileNotFoundError:
            self.query_one('#example').update(TEXT_DATACOMPARISON_FILE_NOT_FIND)
        self.loading_indicator.visible = False

if __name__ == "__main__":
    app = HeaderApp()
    app.run()
