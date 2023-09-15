import docx
from docx.enum.text import WD_COLOR_INDEX

def HighlightCounter(doc, color):
    words = []
    words_arrey = []
    words_count = 0
    for paragraph in doc.paragraphs:
        highlight = ""
        x = False
        for run in paragraph.runs:
            if x == False and highlight != "" and highlight[-1] != " ":
                highlight += " "
            if run.font.highlight_color ==  color:
                highlight += run.text
                x = True
            else:
                x = False
        if highlight:
            words.append(highlight)
    for i in range(len(words)):
        words_arrey += words[i].split(" ")
    words_arrey = list(filter(None, words_arrey))
    print(words_arrey)
    words_count = len(words_arrey)
    return words_count

def AllWordsCounter(doc):
    words = []
    words_arrey = []
    words_count = 0
    for paragraph in doc.paragraphs:
        word = ""
        for run in paragraph.runs:
            word += run.text
        if word:
            words.append(word)
    for i in range(len(words)):
        words_arrey += words[i].split(" ")
    words_count = len(words_arrey)
    return words_count