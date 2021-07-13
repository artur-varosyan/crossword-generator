import random
import sqlite3

blacklist = []
MIN_WORD_LENGTH = 3
MIN_FIRST_WORD_LENGTH = 8


def generate_random_word(word_list):
    try:
        with sqlite3.connect("../resources/Words.db") as connection:
            cursor = connection.cursor()
            cleared = False
            while not cleared:
                rand_num = random.randint(1, 3001)
                query = "SELECT word FROM Words WHERE id = ?"
                try:
                    cursor.execute(query, (rand_num,))
                except sqlite3.DatabaseError:
                    raise sqlite3.DatabaseError("Error while trying to access a random word from the database!")
                row = cursor.fetchall()
                temp_word = row[0][0]
                cleared = check_if_cleared(temp_word, word_list)
    except EnvironmentError:
        raise EnvironmentError("Problem accessing the Words.db database!")

    return temp_word


def check_if_cleared(temp_word, word_list):
    if len(temp_word) < MIN_WORD_LENGTH or temp_word in word_list or temp_word in blacklist:
        cleared = False  # conditions not met
    elif len(word_list) == 0:
        if len(temp_word) > MIN_FIRST_WORD_LENGTH:  # first word in the crossword should be long
            cleared = True
    else:
        cleared = True
    return cleared
