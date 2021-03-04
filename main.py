# Control of performance discipline

from docx import Document


file = open('events.docx', 'rb')
doc = Document(file)
file.close()

'''Поиск по исполнителю'''
for i in range(1, len(doc.tables)):
    for j in range(1, len(doc.tables[i].rows)):
        print(doc.tables[i].rows[j].cells[3].text)

'''Поиск по контролю'''
for i in range(1, len(doc.tables)):
    for j in range(1, len(doc.tables[i].rows)):
        print(doc.tables[i].rows[j].cells[4].text)

'''Поиск по сроку исполнения'''
for i in range(1, len(doc.tables)):
    for j in range(1, len(doc.tables[i].rows)):
        print(doc.tables[i].rows[j].cells[2].text)


