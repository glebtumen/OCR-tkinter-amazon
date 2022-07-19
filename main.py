import boto3
from googletrans import Translator
from tkinter import *
from tkinter import filedialog as fd
import aspose.words as aw

# подключаемся к амазон
client = boto3.client('textract', aws_access_key_id='amazon key',
                      aws_secret_access_key='amazon key', region_name='us-east-2')
# объект переводчика,
translator = Translator()

# объект библиотеки для создания ворда

doc = aw.Document()
builder = aw.DocumentBuilder(doc)

# настройка шрифта
font = builder.font
font.size = 14
font.bold = False
font.name = "Times New Roman"

# настройка параграфа
paragraphFormat = builder.paragraph_format
paragraphFormat.first_line_indent = 8
paragraphFormat.alignment = aw.ParagraphAlignment.JUSTIFY
paragraphFormat.keep_together = True


# функция чтобы открыть файл, считать текст, перевести и вставить в ткинтер
def insert_text():
    file_name = fd.askopenfilenames()  # какой файл открыть

    list_file = list(file_name)

    for file_ in list_file:

        with open(file_, 'rb') as f:  # прочитать в бинарном формате
            file = f.read()

        response = client.detect_document_text(  # отправляем на амазон
            Document={
                'Bytes': file,
            }
        )

        extractedText = []
        for block in response['Blocks']:
            if block['BlockType'] == 'LINE':
                extractedText.append(block['Text'])  # парсим результат

        response = '\n'.join(extractedText)  # соединяем список в одну строку
        # response_rus = translator.translate(response, dest='ru').text #переводим на русский
        text.insert(1.0, response)  # вставляем текст в окно
        text.insert(1.0, '\n')
        text.insert(1.0, '\n')


def extract_text():
    text_to_translate = text.get(1.0, END)
    builder.writeln(text_to_translate)
    file_name = fd.asksaveasfilename(  # спрашиваем как сохранить
        defaultextension='.txt',
        filetypes=(("Word", "*.docx"),
                   ("Text files", "*.txt"),
                   ("All files", "*.*"),
                   ))

    doc.save(file_name)


def translate_to():
    text_to_translate = text.get(1.0, END)

    response = translator.detect(text_to_translate)

    if response.lang == 'en':
        text.delete(1.0, END)

        response_rus = translator.translate(text_to_translate, dest='ru').text  # переводим на русский
        text.insert(1.0, response_rus)  # вставляем текст в окно
        text.insert(1.0, '\n')
        text.insert(1.0, '\n')


    elif response.lang == 'ru':
        text.delete(1.0, END)

        response_rus = translator.translate(text_to_translate, dest='en').text  # переводим на русский
        text.insert(1.0, response_rus)  # вставляем текст в окно
        text.insert(1.0, '\n')
        text.insert(1.0, '\n')


# создание окна ткинтер
root = Tk()
root.geometry('700x500')
text = Text(width=85, height=29)
text.grid(columnspan=2)

b1 = Button(text="Открыть", command=insert_text)
b1.grid(row=1, column=0, sticky=E)

b2 = Button(text="Сохранить", command=extract_text)
b2.grid(row=1, column=1, sticky=W)

b3 = Button(text="Перевести", command=translate_to)
b3.grid(row=1, column=1)

# запуск ткинтер
root.mainloop()
