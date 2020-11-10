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


def find_and_append_tags_from_strs():
    list_of_tags_from_strs = []
    for paragraph in doc.paragraphs:
        tags = how_many_tags(paragraph.text)
        paragraph_text = paragraph.text
        while tags != 0:    # проверка что все теги из строки скопированы
            start = paragraph_text.find('{{')   # символы начала тега
            end = paragraph_text.find('}}', start)  # символы конца тега
            if start != -1 and end != -1 and paragraph_text != '':  # условия, чтобы пустые строки не попали в список
                list_of_tags_from_strs.append(paragraph_text[start + 2:end])
            paragraph_text = paragraph_text[end + 2:]
            # переприсваивание строки, обрезая уже пойманный тег, чтобы ловить несколько тегов в строке,
            # иначе будет ловиться только первый тег в строке
            tags -= 1 # проверка что все теги из строки скопированы
    return list_of_tags_from_strs


def find_and_append_tags_from_tables():
    list_of_tags_from_tables = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                tags = how_many_tags(cell.text)
                cell_text = cell.text
                while tags != 0:    # проверка что все теги из строки скопированы
                    start = cell_text.find('{{')    # символы начала тега
                    end = cell_text.find('}}', start)   # символы конца тега
                    if start != -1 and end != -1 and cell_text != '':  # условия, чтобы пустые строки не попали в список
                        list_of_tags_from_tables.append(cell_text[start + 2:end])
                    cell_text = cell_text[end + 2:]
                    tags -= 1   # проверка что все теги из строки скопированы
    return list_of_tags_from_tables


def making_dict_and_adding_keys(tags_from_strs, tags_from_tables):
    list_of_tags = list(set(tags_from_strs + tags_from_tables))     # суммируем списки, потом делаем из них множество,
    dict_of_tags = {}                                               # чтобы избавиться от повторения, потом снова
    for i in range(0, len(list_of_tags)):                           # приводим к списку
        dict_of_tags[list_of_tags[i]] = 'value of ' + list_of_tags[i]   # присваиваем тегам значение 'value of {{tag}}'
    for value in dict_of_tags.items():
        new_value = input('Введите значение тега "{}": '.format(value[0]))  # присваиваем тегам свои значения
        if new_value != '':     # условие, чтобы пустые строки не попали в словарь
            dict_of_tags[value[0]] = new_value
    return dict_of_tags


def rewriting_tags_to_doc(dict):
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            for value in dict.items():
                if value[0] in run.text:    # run - послед-ность символов в едином форматировании символов
                    start = run.text.find('{{' + value[0])
                    end = run.text.find(value[0] + '}}', start)
                    run.text = run.text[:start] + value[1] + run.text[end + 2 + len(value[0]):]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        for value in dict.items():
                            if value[0] in run.text:    # Без run текст вставляется в без форматирования, которое
                                start = run.text.find('{{' + value[0]) # использовалось в документе
                                end = run.text.find(value[0] + '}}', start)
                                run.text = run.text[:start] + value[1] + run.text[end + 2 + len(value[0]):]


d = making_dict_and_adding_keys(find_and_append_tags_from_strs(), find_and_append_tags_from_tables())
rewriting_tags_to_doc(d)
doc.save('test_updated.docx')

