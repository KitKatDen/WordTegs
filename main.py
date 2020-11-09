from docx import Document

doc = Document('test.docx')


def how_many_tags(paragraph_text):
    n = 0
    char = 0
    while char != len(paragraph_text):
        if paragraph_text[char] == '{' and paragraph_text[char + 1] == '{':
            n += 1
        char += 1
    return n


def find_and_append_teg_from_strs(paragraph):
    tags = how_many_tags(paragraph.text)
    paragraph_text = paragraph.text
    global l
    while tags != 0:
        start = paragraph_text.find('{{')
        end = paragraph_text.find('}}', start)
        if start != -1 and end != -1 and paragraph_text != '':
            # print(paragraph_text[start + 2:end])
            l.append(paragraph_text[start + 2:end])
        paragraph_text = paragraph_text[end + 2:]
        tags -= 1


l = []

for paragraph in doc.paragraphs:
    find_and_append_teg_from_strs(paragraph)
# print(l, len(l))
s = list(set(l))
# print(s, len(s))
d = {}
for i in range(0, len(s)):
    d[s[i]] = 'value of ' + s[i]
print(d, len(d))

for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            print(cell.text)
# print('*' * 20)
# print(doc.tables[2].cell(0, 1).text)


# def find_and_append_teg_from_tables()
# doc.save('test_updated.docx')
