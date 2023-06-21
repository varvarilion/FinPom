from openpyxl import *
from tkinter import *

wb = Workbook()


sheet = wb.active

def excel():
    sheet.column_dimensions['A'].width = 30
    sheet.column_dimensions['B'].width = 10



    sheet.cell(row=1, column=1).value = "Имя"
    sheet.cell(row=1, column=2).value = "Пароль"


def focus1(event):
    name_field.focus_set()

def focus2(event):
    course_field.focus_set()

def clear():
    # clear the content of text entry box
    name_field.delete(0, END)
    course_field.delete(0, END)




def insert():

    if (name_field.get() == "" and
            course_field.get() == "") :

        print("empty input")

    else:


        current_row = sheet.max_row
        current_column = sheet.max_column

        sheet.cell(row=current_row + 1, column=1).value = name_field.get()
        sheet.cell(row=current_row + 1, column=2).value = course_field.get()

        # save the file
        wb.save('balances.xlsx')

        # set focus on the name_field box
        name_field.focus_set()

        # call the clear() function
        clear()

if __name__ == "__main__":
    # create a GUI window
    root = Tk()

    # set the background colour of GUI window
    root.configure(background='light green')

    # set the title of GUI window
    root.title("registration form")

    # set the configuration of GUI window
    root.geometry("500x300")

    excel()

    # create a Form label
    heading = Label(root, text="Form", bg="light green")

    # create a Name label
    name = Label(root, text="Имя", bg="light green")

    # create a Course label
    course = Label(root, text="Пароль", bg="light green")

    heading.grid(row=0, column=1)
    name.grid(row=1, column=0)
    course.grid(row=2, column=0)


    name_field = Entry(root)
    course_field = Entry(root)


    name_field.bind("<Return>", focus1)

    course_field.bind("<Return>", focus2)

    name_field.grid(row=1, column=1, ipadx="100")
    course_field.grid(row=2, column=1, ipadx="100")
    excel()
    submit = Button(root, text="Submit", fg="Black",
                    bg="Red", command=insert)
    submit.grid(row=8, column=1)
    root.mainloop()
wb.save('balances.xlsx')
