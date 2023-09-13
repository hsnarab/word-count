import docx

doc = docx.Document("./test.docx")
for i in doc.paragraphs:
    for j in i.runs:
        print(j.font.highlight_color)