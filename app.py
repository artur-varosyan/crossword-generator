import openpyxl
import shutil
import os
from random_word import random_word_generator
from crossword_allocator import allocate_word
from classes import Word
from crossword_formatting import format_crossword

dev_mode = False

print("Welcome to the Crossword Generator Project written in Python!")
print("It uses open source packages BeautifulSoup and openpyxl")
print("")
input("Press enter to continue")

word_db = openpyxl.load_workbook("WordsList.xlsx")
word_sheet = word_db["Sheet"]

crossword_db = openpyxl.load_workbook("CrosswordTemplate.xlsx")
crossword = crossword_db["Crossword"]
hints = crossword_db["Hints"]

word_list = []
temp_list = []


finished = False
x = 0
while not finished:
    temp_word = random_word_generator(word_sheet, word_list, temp_list)
    allocate_word(temp_word, word_list, temp_list, crossword, hints)
    x = x + 1
    if len(word_list) == 50:
        finished = True
    if dev_mode == True:
        crossword_db.save(f"Crossword{x}.xlsx")

answers = crossword_db.copy_worksheet(crossword)
answers.title = "Answers"
format_crossword(crossword, hints, answers)

crossword_db.save("Crossword.xlsx")
print("Program finished successfully")
print(f"There are {len(word_list)} words in the crossword")
