# Control of performance discipline

from docx import Document
from datetime import date, timedelta, datetime


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
    'Васильцов Д.Г.':'tch10_def@gomel.rw',
    'ТЧЗ-1':'tch10_z1@gomel.rw',
    'Рагина С.М.':'tch10_dom@gomel.rw',
    'Кучеров М.Н.':'tch10_ot@gomel.rw',
    'Дорошенко П.М.':'tch10_nk@gomel.rw',
    'Секретарь':'tch10_tch@gomel.rw'
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

'''Поиск по дате исполнения и повторная рассылка приорететных задач'''
for i in range(1, len(doc.tables)):
    for j in range(1, len(doc.tables[i].rows)):
        day = datetime.strptime(doc.tables[i].rows[j].cells[2].text, '%d.%m.%Y').date()
        date_today = date.today()
        new_date = date_today
        new_date1 = date_today + timedelta(days=3)
        new_date2 = date_today + timedelta(days=1)
        if day == new_date1:
            plan = []
            for text in range(5):
                plan.append(doc.tables[i].rows[j].cells[text].text)
            for key, value in contacts.items():
                if plan[3] == key:
                    print(f'Отправить {plan} на почту {value}')
        elif day == new_date2:
            plan = []
            for text in range(5):
                plan.append(doc.tables[i].rows[j].cells[text].text)
            for key, value in contacts.items():
                if plan[3] == key:
                    print(f'Отправить {plan} на почту {value}')
            for key, value in contacts.items():
                if plan[4] == key:
                    print(f'Отправить {plan} на почту {value}')
        elif day < new_date:
            plan = []
            for text in range(5):
                plan.append(doc.tables[i].rows[j].cells[text].text)
            print(f'Отправить {plan} на почту {contacts.get("Секретарь")}')








# '''Поиск по столбцам'''
# def column_search(column):
#     for i in range(1, len(doc.tables)):
#         for j in range(1, len(doc.tables[i].rows)):
#             print(doc.tables[i].rows[j].cells[column].text)


