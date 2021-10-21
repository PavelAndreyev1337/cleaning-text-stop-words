from typing import Dict
from text_processing.text_cleaner import TextCleaner


class Application:
    def __init__(self):
        self.__text_cleaner = TextCleaner()

    def __print_paragraphs(self, paragraphs: Dict[int, int]):
        for paragraph_number, count in paragraphs.items():
                print("Paragraph: ", paragraph_number + 1, "\n", count)

    def __print_paragraphs_words_frequency(self, paragraph_words_frequency: Dict[str, Dict[int, int]]):
        for word, paragraphs in paragraph_words_frequency.items():
            print("Word: ", word)
            self.__print_paragraphs(paragraphs)

    def run(self):
        self.__text_cleaner.clear()
        print("\nSet of all stop words:", self.__text_cleaner.stop_words)
        print("\nSet of used stop words:", set(
            self.__text_cleaner.last_used_stop_words))
        print("\nNumber of stop words used:", len(
            self.__text_cleaner.last_used_stop_words))
        print("\nPercentage of stop words in the given text:",
              str(int(round(self.__text_cleaner
                            .get_last_stop_words_percentage(), 2) * 100)) + "%")
        self.__text_cleaner.export_last_paragraphs_words_frequency_to_xlsx()
        print("\nFrequent words:", self.__text_cleaner.frequent_words.keys())
        print("\nFrequency of words in paragraphs:")
        self.__print_paragraphs_words_frequency(self.__text_cleaner.paragraph_words_frequency)
        print("\nNumber of words per paragraph:")
        self.__print_paragraphs(self.__text_cleaner.paragraph_words_counts)
        print("\nWords count:", self.__text_cleaner.get_words_count())
        print("\n Relative_frequency:")
        self.__print_paragraphs_words_frequency(self.__text_cleaner.relative_frequency)
