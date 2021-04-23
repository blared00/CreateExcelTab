from tkinter import *
from tkinter import messagebox as mb
from tkinter import filedialog as fd
from project.Sorting import create_itog


num = 5
extra = 0

def create_entry(n):
    text.append(Entry(width=50))
    text[n].grid(row=n+1, column=0, sticky=W)
    b.append(Button(text="...", command=lambda :insert_text()))
    b[n].grid(row=n+1, column=1, sticky=W)
    number.append(Entry(width=10))
    number[n].grid(row=n+1, column=2, sticky=W)
    def insert_text():
        file_name = fd.askopenfilename(filetypes=(("EXCEL files", "*.xls;*.xlsx"),
                                                  ))
        if not text[n].get():
            text[n].insert(0, file_name)
        else:
            text[n].delete(0, END)
            text[n].insert(0, file_name)


def add_extra():
    pass


def read_files_path():
    files_df = []
    for n in range(num + extra):
        if not text[n].get() or not number[n].get():
            continue
        files_df.append({'sourse': text[n].get().replace('/', '\\'), 'num_yach': int(number[n].get())})

    return files_df



root = Tk()
label = Label(text='Выберете Excel-фаилы',
              font=("Times New Romans",
                    24, "bold")).grid(columnspan=3, sticky=W+E)
number = []
text = []
b = []


for n in range(num):
    create_entry(n)
b_create = Button(text="Создать сводный Excel-фаил", command=lambda :create_itog(read_files_path()))
b_create.grid(row=num+extra+1, columnspan=3, sticky=W+E)

root.mainloop()
