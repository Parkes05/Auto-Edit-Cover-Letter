from tkinter import *
from tkinter.tix import *
from docx import Document
from docx2pdf import convert

DICT = {}
PATH = 'C:/Users/benny/Desktop/Cover_Letter_Benjamin_Muoka'

def create_dictionary(name, date, address, type, position):
    global DICT
    DICT = {
        '{company}': name,
        '{date}': date,
        '{addre}': address[0],
        '{addres}': address[1],
        '{address}': address[2],
        '{type}': type,
        '{position}': position,
        '{uposition}': position.upper(),
        }

def replace_words_and_save():
    with open('Cover_Letter_Benjamin_Muoka.docx', 'r+b') as f:
        doc = Document(f)
        for key in DICT:
            for idxPara, elemPara in enumerate(doc.paragraphs):
                for run in elemPara.runs:
                    if key in run.text:
                        run.text = DICT[key]
        doc.save(f'{PATH}.docx')

def convert_to_pdf():
    convert(f'{PATH}.docx', f'{PATH}.pdf')

def gui_for_input():
    root = Tk()
    root.title('cover_letter_update')
    tool_tip = Balloon(root)

    label1 = Label(root, text='Company Name')
    label1.grid(column=1, row=1, sticky='w')
    entry1 = Entry(root, width=20)
    entry1.grid(column=2, row=1)

    label2 = Label(root, text='Date')
    label2.grid(column=1, row=2, sticky='w')
    entry2 = Entry(root, width=20)
    entry2.grid(column=2, row=2)
    btn1 = Button(root, text='?', padx=5, borderwidth=0)
    btn1.grid(column=3, row=2, padx=5)

    label3 = Label(root, text='Company Address')
    label3.grid(column=1, row=3, sticky='w')
    entry3 = Entry(root, width=20)
    entry3.grid(column=2, row=3)
    entry4 = Entry(root, width=20)
    entry4.grid(column=2, row=4)
    entry5 = Entry(root, width=20)
    entry5.grid(column=2, row=5)

    label4 = Label(root, text='Company Type')
    label4.grid(column=1, row=6, sticky='w')
    entry6 = Entry(root, width=20)
    entry6.grid(column=2, row=6)
    btn2 = Button(root, text='?', padx=5, borderwidth=0)
    btn2.grid(column=3, row=6, padx=5)

    label5 = Label(root, text='Job Position')
    label5.grid(column=1, row=7, sticky='w')
    entry7 = Entry(root, width=20)
    entry7.grid(column=2, row=7)

    def buttoncommand():
        name = entry1.get()
        date = entry2.get()
        address = [entry3.get(), entry4.get(), entry5.get()]
        type = entry6.get()
        position = entry7.get()
        create_dictionary(name, date, address, type, position)
        replace_words_and_save()
        convert_to_pdf()
        root.destroy()

    btn2 = Button(root, text='confirm', background='blue', command=buttoncommand, foreground='white')
    btn2.grid(columnspan=3, row=8, pady=10)

    tool_tip.bind_widget(btn1, balloonmsg="Month Day, Year")
    tool_tip.bind_widget(btn2, balloonmsg="Institution, Company or University")

    root.mainloop()


gui_for_input()
