import requests


def find_definition(word):
    url = f"https://api.dictionaryapi.dev/api/v2/entries/en_GB/{word}"
    response = requests.get(url)
    if response.status_code == 404:  # if GB dictionary does not contain word try US dictionary
        url = f"https://api.dictionaryapi.dev/api/v2/entries/en_US/{word}"
        response = requests.get(url)
    if response.status_code == 200:  # success
        output = response.json()
        meanings = output[0]["meanings"]
        first_meaning = meanings[0]
        definitions = first_meaning["definitions"]
        definition = definitions[0]["definition"]
        print(definition)
        return definition
    else:
        print("Incorrect error code")
        print(f"API Request call returned: {response.status_code}. Word queried: {word}")
        print("Returning the word as definition")
        return f"Definition not found, answer: {word}"
