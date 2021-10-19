from nltk import download
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from docx import Document
from typing import List, Set


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

    def clear(self):
        self.__reset_results()
        self.__clear_paragraphs()
        self.__clear_tables()
        self.__document.save(self.__output_file_path)
