from nltk import download
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from docx import Document
from typing import Dict, List, Set
from openpyxl import Workbook
from collections import Counter


class TextCleaner:
    def __init__(self, language: str = "russian",
                 file_path: str = "text.docx",
                 output_file_path: str = "Андреєв_етап1.docx"):
        download('stopwords')
        download('punkt')
        self.__language = language
        self.__stop_words = set(stopwords.words(self.__language))
        self.__file_path = file_path
        self.__output_file_path = output_file_path
        self.__document = Document(self.__file_path)
        self.__workbook = None
        # used pr-cy.ru {key:word : value:the root of the word, }
        self.__frequent_words = []
        # the frequency of words in the text {key: word, value: count,}
        self.__text_words_frequency = {}
        # frequency of words in paragraphs {key: word, value: {key:paragraph_number: value: count,},}
        self.__paragraph_words_frequency = {}
        # the number of words in a paragraph {key: paragraph_number, value: count}
        self.__paragraph_words_counts = {}
        self.__frequent_words_count = 0
        self.__relative_frequency = {}  # 3-digit accuracy after coma
        self.__output_xlsx_file_path = "output.xlsx"
        self.__last_text_words_count = 0
        self.__last_used_stop_words = []

    @property
    def language(self) -> str:
        return self.__language

    @language.setter
    def language(self, language: str) -> None:
        self.__language = language
        self.__stop_words = set(stopwords.words(self.__language))

    @property
    def file_path(self) -> str:
        return self.__file_path

    @property
    def stop_words(self) -> Set[str]:
        return self.__stop_words

    @file_path.setter
    def file_path(self, file_path: str) -> None:
        self.__file_path = file_path
        self.__document = Document(file_path)

    @property
    def last_text_words_count(self) -> int:
        return self.__last_text_words_count

    @property
    def last_used_stop_words(self) -> List[str]:
        return self.__last_used_stop_words

    @property
    def frequent_words(self):
        return self.__frequent_words

    @property
    def text_words_frequency(self):
        return self.__text_words_frequency

    @property
    def paragraph_words_frequency(self):
        return self.__paragraph_words_frequency

    @property
    def paragraph_words_counts(self):
        return self.__paragraph_words_counts

    @property
    def relative_frequency(self):
        return self.__relative_frequency

    def get_words_count(self):
        return sum([words_count for words_count in self.__paragraph_words_counts.values()])

    def get_last_stop_words_percentage(self) -> float:
        return len(self.__last_used_stop_words) / self.__last_text_words_count

    def __clear_runs(self, paragraph):
        for run in paragraph.runs:
            for word in word_tokenize(run.text):
                self.__last_text_words_count += 1
                if word.lower() in self.__stop_words:
                    self.__last_used_stop_words.append(word)
                    run.text = run.text.replace(f" {word} ", " ")

    def __clear_paragraphs(self):
        for paragraph in self.__document.paragraphs:
            self.__clear_runs(paragraph)

    def __clear_tables(self):
        for table in self.__document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self.__clear_runs(paragraph)

    def __reset_results(self):
        self.__last_text_words_count = 0
        self.__last_used_stop_words = []

    def __calculate_words_count(self):
        self.__document = Document(self.__output_file_path)
        for key_word, root in self.__frequent_words.items():
            self.__text_words_frequency[key_word] = 0
            self.__paragraph_words_frequency[key_word] = {}
            for i, paragraph in enumerate(self.__document.paragraphs):
                self.__paragraph_words_frequency[key_word][i] = 0
                self.__paragraph_words_counts[i] = 0
                for run in paragraph.runs:
                    words = word_tokenize(run.text)
                    self.__paragraph_words_counts[i] += len(words)
                    for word in words:
                        if word.startswith(root):
                            self.__text_words_frequency[key_word] += 1
                            self.__paragraph_words_frequency[key_word][i] += 1
        self.__workbook.save(self.__output_xlsx_file_path)
        self.__text_words_frequency = {word: count for word, count in
                                       sorted(self.__text_words_frequency.items(),
                                              key=lambda item: item[1],
                                              reverse=True)}

    def __calculate_relative_frequency(self):
        for word, paragraphs in self.__paragraph_words_frequency.items():
            self.__relative_frequency[word] = {}
            for paragraph_number, words_count in paragraphs.items():
                paragraph_words_count = self.__paragraph_words_counts[paragraph_number]
                if not paragraph_words_count:
                    self.__relative_frequency[word][paragraph_number] = 0
                else:
                    self.__relative_frequency[word][paragraph_number] = round(
                        words_count / self.__paragraph_words_counts[paragraph_number], 3)

    def export_last_relative_frequency_to_csv(self,
                                              frequent_words: Dict[str, str] = {},
                                              output_xlsx_file_path: str = "Андреєв_етап3.xlsx",
                                              frequent_words_count: int = 21):  # stage 3
        self.__output_xlsx_file_path = output_xlsx_file_path
        self.__frequent_words_count = frequent_words_count
        self.__workbook = Workbook()
        if not frequent_words:
            self.__frequent_words = {  # result from pr-cy.ru
                "автосамосвал": "автосамосвал",
                "модель": "модел",
                "транспортный": "транспортн",
                "движение": "движен",
                "карьер": "карьер",

                "работа": "работ",
                "состояние": "состоян",
                "разгрузка": "разгрузк",
                "блок": "блок",
                "пункт": "пункт",

                "погрузка": "погрузк",
                "система": "систем",
                "экскаватор": "экскаватор",
                "время": "врем",
                "управление": "управлен",

                "имитационный": "имитационн",
                "параметр": "параметр",
                "временить": "времен",
                "скорость": "скорост",
                "цикл": "цикл",

                "граф": "граф",
            }
        else:
            self.__frequent_words = frequent_words
        self.__calculate_words_count()
        self.__calculate_relative_frequency()

    def clear(self):
        self.__reset_results()
        self.__clear_paragraphs()
        self.__clear_tables()
        self.__document.save(self.__output_file_path)
