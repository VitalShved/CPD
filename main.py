# Control of performance discipline

from docx import Document


contacts = {
    'Хоменок Ю.В.':'tch10_meh@gomel.rw',
    'Королев С.Н.':'tch10_electr@gomel.rw',
    'Зятиков А.А.':'tch10_snab@gomel.rw',
    'ТЧГ':'tch10_gi@gomel.rw',
    'Швед В.А.':'tch10_tcht@gomel.rw',
    'Ткачев А.С.':'tch10_hoz@gomel.rw',
    'Говор П.В.':'tch10_rmu@gomel.rw',
    'Зезюлин П.В.':'tch10_to@gomel.rw',
    'Говязо Е.А.':'tch10_klad@gomel.rw',
    'Васильцов Г.И.':'tch10_def@gomel.rw',
    'ТЧЗ-1':'tch10_z1@gomel.rw',
    'Рагина С.М.':'tch10_dom@gomel.rw',
    'Кучеров М.Н.':'tch10_ot@gomel.rw',
    'Дорошенко П.М.':'tch10_nk@gomel.rw'
}

file = open('events.docx', 'rb')
doc = Document(file)
file.close()

newsletter = set()

'''Поиск по исполнителю'''
for i in range(1, len(doc.tables)):
    for j in range(1, len(doc.tables[i].rows)):
        newsletter.add(doc.tables[i].rows[j].cells[3].text)

'''Поиск по контролирующему'''
for i in range(1, len(doc.tables)):
    for j in range(1, len(doc.tables[i].rows)):
        newsletter.add(doc.tables[i].rows[j].cells[4].text)

'''Рассылка документа причастным'''
for name in newsletter:
    for key, value in contacts.items():
        if name == key:
            print(f'Отправить файл {doc} на почту {value}')


# '''Поиск по контролирующему'''
# for i in range(1, len(doc.tables)):
#     for j in range(1, len(doc.tables[i].rows)):
#         print(doc.tables[i].rows[j].cells[4].text)
#
# '''Поиск по сроку исполнения'''
# for i in range(1, len(doc.tables)):
#     for j in range(1, len(doc.tables[i].rows)):
#         print(doc.tables[i].rows[j].cells[2].text)

# '''Поиск по столбцам'''
# def column_search(column):
#     for i in range(1, len(doc.tables)):
#         for j in range(1, len(doc.tables[i].rows)):
#             print(doc.tables[i].rows[j].cells[column].text)


