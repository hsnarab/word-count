import docx
from docx.enum.text import WD_COLOR_INDEX

def HighlightCounter(doc, color):
    special_chars = ["!", "\"", "@", "#", "$", "%", "^", "&", "*", "(", ")", "-", "_", "+", "=", "{", "}", ":", ";", "'", "\\", "|", "~", "`", ",", "<", ".", ">", "/", "?", "،", "؛", ",", "]", "[", "«", "»", ":", "ء", "؟"]
    words_count = 0
    for para in doc.paragraphs:
        for part in para.runs:
            if part.font.highlight_color ==  color:
                    words = part.text
                    words_arrey = words.split(" ")
                    if words_arrey[0] == "" or words_arrey[-1] == "":
                        words_arrey.remove("")
                    if words_arrey[-1] == "":
                        words_arrey.remove("")
                    if words_arrey != 0:
                        for schar in special_chars:
                            for char in words_arrey:
                                if schar == char:
                                    words_arrey.remove(char)
                    if len(words_arrey) == 0:
                        words_arrey = 0
                    if words_arrey != 0:
                        words_count += len(words_arrey)
    return words_count

def AllWordsCounter(doc):
    special_chars = ["!", "\"", "@", "#", "$", "%", "^", "&", "*", "(", ")", "-", "_", "+", "=", "{", "}", ":", ";", "'", "\\", "|", "~", "`", ",", "<", ".", ">", "/", "?", "،", "؛", ",", "]", "[", "«", "»", ":", "ء", "؟"]
    words_count = 0
    for para in doc.paragraphs:
        for part in para.runs:
            words = part.text
            words_arrey = words.split(" ")
            if words_arrey[0] == "" or words_arrey[-1] == "":
                words_arrey.remove("")
            if words_arrey[-1] == "":
                words_arrey.remove("")
            if words_arrey != 0:
                for schar in special_chars:
                    for char in words_arrey:
                        if schar == char:
                            words_arrey.remove(char)
            if len(words_arrey) == 0:
                words_arrey = 0
            if words_arrey != 0:
                words_count += len(words_arrey)
    return words_count