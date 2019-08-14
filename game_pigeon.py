from itertools import combinations, permutations
from openpyxl import load_workbook


class GamePigeon(object):

    @staticmethod
    def gen_words():

        wb = load_workbook(filename='all_words.xlsx')
        ws = wb.active
        all_words = set()

        for row in ws.values:
            if row[0] is None:
                break
            all_words.add(row[0])

        return all_words

    @staticmethod
    def check_word(word, input_set):
        if word in input_set:
            return True
        else:
            return False

    def word_hunt(self, letters):

        if len(letters) != 16:
            return 'error: invalid input'

        letters = letters.upper()
        letters_array = []

        for letter in letters:
            letters_array.append(letter)

        moves_dict = {
            1: [2, 5, 6],
            2: [1, 3, 5, 6, 7],
            3: [2, 4, 6, 7, 8],
            4: [3, 7, 8],
            5: [1, 2, 6, 9, 10],
            6: [1, 2, 3, 5, 7, 9, 10, 11],
            7: [2, 3, 4, 6, 8, 10, 11, 12],
            8: [3, 4, 7, 11, 12],
            9: [5, 6, 10, 13, 14],
            10: [5, 6, 7, 9, 11, 13, 14, 15],
            11: [6, 7, 8, 10, 12, 14, 15, 16],
            12: [7, 8, 11, 15, 16],
            13: [9, 10, 14],
            14: [9, 10, 11, 13, 15],
            15: [10, 11, 12, 14, 16],
            16: [11, 12, 15]
        }

        moves_list = []

        for start_point in range(16):
            start = start_point + 1

            temp_combo = [start]

            for move2 in moves_dict[start]:
                temp_combo.append(move2)
                for move3 in moves_dict[move2]:
                    if move3 not in temp_combo:
                        temp_combo.append(move3)
                        for move4 in moves_dict[move3]:
                            if move4 not in temp_combo:
                                temp_combo.append(move4)
                                moves_list.append(temp_combo[:])
                                for move5 in moves_dict[move4]:
                                    if move5 not in temp_combo:
                                        temp_combo.append(move5)
                                        moves_list.append(temp_combo[:])
                                        for move6 in moves_dict[move5]:
                                            if move6 not in temp_combo:
                                                temp_combo.append(move6)
                                                moves_list.append(temp_combo[:])
                                                for move7 in moves_dict[move6]:
                                                    if move7 not in temp_combo:
                                                        temp_combo.append(move7)
                                                        moves_list.append(temp_combo[:])
                                                        for move8 in moves_dict[move7]:
                                                            if move8 not in temp_combo:
                                                                temp_combo.append(move8)
                                                                moves_list.append(temp_combo[:])
                                                                temp_combo.remove(move8)
                                                        temp_combo.remove(move7)
                                                temp_combo.remove(move6)
                                        temp_combo.remove(move5)
                                temp_combo.remove(move4)
                        temp_combo.remove(move3)
                temp_combo.remove(move2)

        all_words_list = self.gen_words()

        letter_combos = []

        for move in moves_list:
            word = ''
            for position in move:
                word = word + letters_array[position - 1]
            letter_combos.append(word)

        letter_combos = list(dict.fromkeys(letter_combos))

        final_list = []

        for word in letter_combos:
            if self.check_word(word, all_words_list):
                final_list.append(word)

        final_list.sort(key=lambda s: len(s))
        final_list = final_list[::-1]
        return final_list

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
    return test_game.word_hunt(test_string)


print(test_class('siprnhnctaeilseo'))
