# Control of performance discipline


from docx import Document


file = open('events.docx', 'rb')
document = Document(file)
print(document)
file.close()

