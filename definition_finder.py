from bs4 import BeautifulSoup
import requests

def find_definition(word):
    print("Started webscraping...")
    URL = f"https://www.dictionary.com/browse/{word}"
    page = requests.get(URL)

    soup = BeautifulSoup(page.text, "html.parser")

    definition = soup.find(class_="one-click-content css-ana4le-PosSupportingInfo e1q3nk1v1").get_text()
    output = ""

    for char in definition:
        if char == ";" or char == ".":
            return output
        else:
            output += char
    return output