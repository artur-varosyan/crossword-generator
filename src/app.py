import openpyxl
from random_word import generate_random_word
from crossword_allocator import allocate_word
from crossword_formatting import format_crossword


def main():
    dev_mode = False

    print("Welcome to the Crossword Generator Project written in Python!")
    print("")
    input("Press enter to continue")
    print("Creating crossword...")

    crossword_db = openpyxl.load_workbook("../resources/CrosswordTemplate.xlsx")
    crossword = crossword_db["Crossword"]
    hints = crossword_db["Hints"]

    word_list = []

    finished = False
    x = 0
    while not finished:
        temp_word = generate_random_word(word_list)
        allocate_word(temp_word, word_list, crossword, hints)
        x = x + 1
        if len(word_list) == 50:
            finished = True
        if dev_mode:
            crossword_db.save(f"Crossword{x}.xlsx")

    answers = crossword_db.copy_worksheet(crossword)
    answers.title = "Answers"
    format_crossword(crossword, answers)

    crossword_db.save("../Crossword.xlsx")
    print("")
    print("Program finished successfully")
    print(f"There are {len(word_list)} words in the crossword")


if __name__ == '__main__':
    main()
