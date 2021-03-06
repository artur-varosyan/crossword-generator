import random
import openpyxl

blacklist = []

def random_word_generator(word_sheet, word_list, temp_list):
    cleared = False
    while not cleared:
        id = random.randint(1, 3000)
        temp_cell = word_sheet.cell(id + 2, 3)
        temp_word = temp_cell.value
        print(temp_word)
        if len(temp_word) < 3 or temp_word in word_list or temp_word in blacklist:
            cleared = False
            print("Conditions not met")
        elif len(word_list) == 0:
            # print("It is the first word")
            if len(temp_word) > 8:
                # print("long enough for first word")
                cleared = True
        else:
            cleared = True
            print("Conditions met. Word is returned")
    return temp_word
