import openpyxl
from openpyxl.comments import Comment

from classes import Word
from definition_finder import find_definition
import random


def allocate_word(temp_word, word_list, temp_list, crossword, hints):
    if len(word_list) == 0:
        left_indent = (25 - len(temp_word)) / 2
        left_indent = round(left_indent)
        word_list.append(Word(temp_word, "horizontal", 12, left_indent, len(temp_word)))

        map_word(word_list, crossword)

        definition = find_definition(word_list[-1].name)
            
        hints.cell(3+len(word_list), 2).value = word_list[-1].orientation
        hints.cell(3+len(word_list), 3).value = definition

    else:
        matched, word_list = match_words(temp_word, word_list, crossword)
        if matched:
            map_word(word_list, crossword)
            definition = find_definition(word_list[-1].name)
            
            hints.cell(3+len(word_list), 2).value = word_list[-1].orientation
            hints.cell(3+len(word_list), 3).value = definition

        else:
            print("New random word needed")
            # get new word


def match_words(temp_word, word_list, crossword):
    i = 0 # 0 represents the previous word, 1 represents the word before that one etc
    for words in word_list:
        comparison_word = word_list[len(word_list) - 1 - i].name
        print(f"Now trying... {comparison_word}")
        index = 0
        index_temp = 0
        matched = False
        for letter in comparison_word:
            for letter_temp in temp_word:
                if letter_temp == letter:
                    print("Matched")
                    print(f"Index of First: {index}")
                    print(f"Index of new: {index_temp}")
                    matched = True
                    overlap, word_list = check_surroundings(temp_word, word_list, index, index_temp, crossword, i)
                    if not overlap:
                        return matched, word_list
                index_temp = index_temp + 1
            index_temp = 0
            index = index + 1
        print("No match for this comparison word!")
        i = i + 1

    print("No match AT ALL in crossword")
    matched = False
    return matched, word_list


def check_surroundings(temp_word, word_list, index, index_temp, crossword, i):

    # STEP 1 + STEP 2
    if word_list[len(word_list) - 1 - i].orientation == "horizontal":
        actual_row = word_list[len(word_list) - 1 - i].starting_row
        actual_column = word_list[len(word_list) - 1 - i].starting_column + index
        new_orientation = "vertical"
    else:
        actual_row = word_list[len(word_list) - 1 - i].starting_row + index
        actual_column = word_list[len(word_list) - 1 - i].starting_column
        new_orientation = "horizontal"
    print(f"new orientation = {new_orientation}")

    print(f"actual row = {actual_row}")
    print(f"actual column = {actual_column}")

    # STEP 3
    before_intersection = index_temp
    after_intersection = len(temp_word) - index_temp - 1

    # STEP 4
    # =  No horizontal letters allowed
    # ! No vertical letters allowed
    # @ no letters at all
    overlap = False

    try:

        # Bug fix below
        if new_orientation == "vertical":
            print("ONE BEFORE First character")
            print(crossword.cell(actual_row - before_intersection - 1, actual_column).value)
            if crossword.cell(actual_row - before_intersection - 1, actual_column).value is not None and crossword.cell(actual_row - before_intersection - 1, actual_column).value != "=" and crossword.cell(actual_row - before_intersection - 1, actual_column).value != "!" and crossword.cell(actual_row - before_intersection - 1, actual_column).value != "@":
                overlap = True
            print("ONE AFTER First character")
            print(crossword.cell(actual_row + after_intersection + 1, actual_column).value)
            if crossword.cell(actual_row + after_intersection + 1, actual_column).value is not None and crossword.cell(actual_row + after_intersection + 1, actual_column).value != "=" and crossword.cell(actual_row + after_intersection + 1, actual_column).value != "!" and crossword.cell(actual_row + after_intersection + 1, actual_column).value != "@":
                overlap = True
        else:
            print("ONE BEFORE First character")
            print(crossword.cell(actual_row, actual_column - before_intersection - 1).value)
            if crossword.cell(actual_row, actual_column - before_intersection - 1).value is not None and crossword.cell(actual_row, actual_column - before_intersection - 1).value != "=" and crossword.cell(actual_row, actual_column - before_intersection - 1).value != "!" and crossword.cell(actual_row, actual_column - before_intersection - 1).value != "@":
                overlap = True
            print("ONE AFTER First character")
            print(crossword.cell(actual_row, actual_column + after_intersection + 1).value)
            if crossword.cell(actual_row, actual_column + after_intersection + 1).value is not None and crossword.cell(actual_row, actual_column + after_intersection + 1).value != "=" and crossword.cell(actual_row, actual_column + after_intersection + 1).value != "!" and crossword.cell(actual_row, actual_column + after_intersection + 1).value != "@":
                overlap = True
        
        print("before intersection")
        for cell in range(0, before_intersection):
            # Bug fix (shouldn't change letters now?) below
            if overlap == True:
                continue
            print(cell)
            if new_orientation == "vertical":
                if (actual_row - cell) > 25 or actual_column > 25:
                        overlap = True
                elif crossword.cell(actual_row - 1 - cell, actual_column).value is not None and crossword.cell(actual_row - 1 - cell, actual_column).value != "=":
                    overlap = True
            else:
                if (actual_column - cell) > 25 or actual_row > 25:
                        overlap = True
                elif crossword.cell(actual_row, actual_column - 1 - cell).value is not None and crossword.cell(actual_row, actual_column - 1 - cell).value != "!":
                    overlap = True
            print(overlap)
        print("after intersection")
        for cell in range(1, after_intersection + 1):
            # Bug fix (shouldn't change letters now?) below
            if overlap == True:
                continue
            print(cell)
            if new_orientation == "vertical":
                if (actual_row + cell) > 25 or actual_column > 25:
                    overlap = True
                elif crossword.cell(actual_row + cell, actual_column).value is not None and crossword.cell(actual_row + cell, actual_column).value != "=":
                    overlap = True
            else:
                if (actual_column + cell) > 25 or actual_row > 25:
                    overlap = True
                if crossword.cell(actual_row, actual_column + cell).value is not None and crossword.cell(actual_row, actual_column + cell).value != "!":
                    overlap = True
    except ValueError:
        print("OUT OF RANGE!")
        overlap = True
    print(overlap)

    # STEP 5
    if not overlap:
        if new_orientation == "vertical":
            word_list.append(
                Word(temp_word, new_orientation,
                     actual_row - before_intersection,
                     actual_column, len(temp_word)))
        else:
            word_list.append(
                Word(temp_word, new_orientation,
                     actual_row,
                     actual_column - before_intersection, len(temp_word)))

    print(f"starting_row = {word_list[len(word_list) - 1].starting_row}")
    print(f"starting_column = {word_list[len(word_list) - 1].starting_column}")
    return overlap, word_list

# EXCEPTIONS ARE NEEDED FOR THE SPACE ALGORITH
# - If there's another character
# - If there's another space

def map_word(word_list, crossword):
    new_word = word_list[len(word_list) - 1]
    print(f"new_word.orientation = {new_word.orientation}")
    if new_word.orientation == "vertical":
        i = 0
        temp_row = new_word.starting_row
        for letter in new_word.name:
            if i == 0:
                # Comments
                i = 1
                crossword.cell(temp_row, new_word.starting_column).value = letter
                comment = Comment(f"{len(word_list)} {new_word.orientation}", "CrosswordGenerator")
                if crossword.cell(temp_row, new_word.starting_column).comment is None:
                    crossword.cell(temp_row, new_word.starting_column).comment = comment
                else:
                    crossword.cell(temp_row, new_word.starting_column).comment.text = crossword.cell(temp_row, new_word.starting_column).comment.text + " " + comment.text
            else:
                crossword.cell(temp_row, new_word.starting_column).value = letter

            if temp_row == new_word.starting_row:
                map_limits(temp_row - 1, new_word.starting_column, "both", crossword)
            elif temp_row == (new_word.starting_row + new_word.length - 1):
                map_limits(temp_row + 1, new_word.starting_column, "both", crossword)

            map_limits(temp_row, new_word.starting_column - 1, "vertical", crossword)
            map_limits(temp_row, new_word.starting_column + 1, "vertical", crossword)

            temp_row = temp_row + 1
    else:
        i = 0
        temp_column = new_word.starting_column
        for letter in new_word.name:
            if i == 0:
                # Comments
                i = 1
                crossword.cell(new_word.starting_row, temp_column).value = letter
                comment = Comment(f"{len(word_list)} {new_word.orientation}", "CrosswordGenerator")
                if crossword.cell(new_word.starting_row, temp_column).comment is None:
                    crossword.cell(new_word.starting_row, temp_column).comment = comment
                else:
                    crossword.cell(new_word.starting_row, temp_column).comment.text = crossword.cell(new_word.starting_row, temp_column).comment.text + " " + comment.text
            else:
                crossword.cell(new_word.starting_row, temp_column).value = letter
            
            if temp_column == new_word.starting_column:
                map_limits(new_word.starting_row, temp_column - 1, "both", crossword)
            elif temp_column == (new_word.starting_column + new_word.length - 1):
                map_limits(new_word.starting_row, temp_column + 1, "both", crossword)
            
            map_limits(new_word.starting_row - 1, temp_column, "horizontal", crossword)
            map_limits(new_word.starting_row + 1, temp_column, "horizontal", crossword)

            temp_column = temp_column + 1


def map_limits(row, column, orientation, crossword):
    print(f"map limits running, row= {row} column={column}")
    print(f"crossword cell = {crossword.cell(row, column).value}")
    if orientation == "vertical":
        if crossword.cell(row, column).value is None:
            crossword.cell(row, column).value = "!"
        elif crossword.cell(row, column).value == "=":
            crossword.cell(row, column).value = "@"
    elif orientation == "horizontal":
        if crossword.cell(row, column).value is None:
            crossword.cell(row, column).value = "="
        elif crossword.cell(row, column).value == "!":
            crossword.cell(row, column).value = "@"
    else:
        if crossword.cell(row, column).value is None:
            crossword.cell(row, column).value = "@"
        elif crossword.cell(row, column).value == "!" or crossword.cell(row, column).value == "=":
            crossword.cell(row, column).value = "@"
    print(f"crossword cell NEW = {crossword.cell(row, column).value}")

