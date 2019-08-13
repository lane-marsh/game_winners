from itertools import combinations, permutations
from openpyxl import load_workbook


class GamePigeon(object):

    @staticmethod
    def gen_words():

        wb = load_workbook(filename='all_words.xlsx')
        ws = wb.active
        all_words_array = []

        for row in ws.values:
            if row[0] is None:
                break
            all_words_array.append(row[0])

        return all_words_array

    @staticmethod
    def check_word(word, array):
        if word in array:
            return True
        else:
            return False

    def anagrams(self, letters):

        letters = letters.upper()
        letters_array = []
        words_array = []

        for letter in letters:
            letters_array.append(letter)

        for each in range(3, len(letters_array) + 1):
            for subset in combinations(letters_array, each):
                for perm in permutations(subset):
                    words_array.append(''.join(perm))

        all_words_list = self.gen_words()

        final_list = []

        for word in words_array:
            if self.check_word(word, all_words_list):
                final_list.append(word)

        final_list = list(dict.fromkeys(final_list))[::-1]

        return final_list


def test_class(test_string):
    test_game = GamePigeon()
    return test_game.anagrams(test_string)


print(test_class('wordss'))
