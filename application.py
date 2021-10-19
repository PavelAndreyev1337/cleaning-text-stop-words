from text_processing.text_cleaner import TextCleaner


class Application:
    def __init__(self):
        self.__text_cleaner = TextCleaner()

    def run(self):
        self.__text_cleaner.clear()
        print("Set of all stop words:", self.__text_cleaner.stop_words)
        print("Set of used stop words:", set(
            self.__text_cleaner.last_used_stop_words))
        print("Number of stop words used:", len(
            self.__text_cleaner.last_used_stop_words))
        print("Percentage of stop words in the given text:",
              str(int(round(self.__text_cleaner.get_last_stop_words_percentage(), 2) * 100)) + "%")
