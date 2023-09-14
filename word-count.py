import docx
from docx.enum.text import WD_COLOR_INDEX

yellow_words = 0
red_words = 0
special_chars = ["!", "\"", "@", "#", "$", "%", "^", "&", "*", "(", ")", "-", "_", "+", "=", "{", "}", ":", ";", "'", "\\", "|", "~", "`", ",", "<", ".", ">", "/", "?", "،", "؛", ",", "]", "[", "«", "»", ":", "ء", "؟"]


doc = docx.Document("./test.docx")
for i in doc.paragraphs:
    for j in i.runs:
        if j.font.highlight_color ==  WD_COLOR_INDEX.YELLOW:
                words_count = 0
                words = j.text
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
                yellow_words += words_count
        if j.font.highlight_color ==  WD_COLOR_INDEX.RED:
                words_count = 0
                words = j.text
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
                red_words += words_count
print("Yellow: " + str(yellow_words))
print("Red: " + str(red_words))