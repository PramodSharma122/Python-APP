from tkinter import *
from tkinter import font
from tkinter.ttk import Combobox
import tkinter as tk
from tkinter import messagebox
import openpyxl, xlrd
from openpyxl import Workbook
import pathlib
from PIL import Image, ImageTk


def create_file():
  file = Workbook()
  sheet = file.active
  sheet['A1'] = 'Full Name'
  sheet['B1'] = "Phone Numbers"
  sheet['C1'] = "Age"
  sheet['D1'] = "Gender"
  sheet['E1'] = "Address"

  file.save('Backend_data.xlsx')


root = Tk()
root.title("Data Entry")
root.geometry('700x400+300+200')
root.resizable(False, False)
root.configure(bg='#326273')

file = pathlib.Path('Backend_data.xlsx')
if not file.exists():
  create_file()


def submit():
  name = nameValue.get()
  contact = contactValue.get()
  age = AgeValue.get()
  gender = gender_combobox.get()
  address = addressEntry.get(1.0, END)

  try:
    file = openpyxl.load_workbook('Backend_data.xlsx')
    sheet = file.active

    sheet.append([name, contact, age, gender, address])
    file.save('Backend_data.xlsx')

    messagebox.showinfo('Success', 'Data Saved')

    nameValue.set("")
    contactValue.set("")
    AgeValue.set("")
    addressEntry.delete(1.0, END)
  except Exception as e:
    messagebox.showerror('Error', str(e))


def clear():
  nameValue.set("")
  contactValue.set("")
  AgeValue.set("")
  addressEntry.delete(1.0, END)


#icon
image = Image.open('logo.jpeg')
icon_image = ImageTk.PhotoImage(image)
root.iconphoto(False, icon_image)

# heading
Label(root,
      text="Please fill in the details below: ",
      font='arial 13',
      bg='#326273',
      fg='#fff').place(x=20, y=20)

# label
Label(root, text='Name : ', font=23, bg="#326273", fg="#fff").place(x=50, y=100)
Label(root, text='Contact No : ', font=23, bg="#326273", fg="#fff").place(x=50,
                                                                       y=150)
Label(root, text='Age : ', font=23, bg="#326273", fg="#fff").place(x=50, y=200)
Label(root, text='Gender : ', font=23, bg="#326273", fg="#fff").place(x=370,
                                                                   y=200)
Label(root, text='Address : ', font=23, bg="#326273", fg="#fff").place(x=50,
                                                                    y=250)

# Entry
nameValue = StringVar()
contactValue = StringVar()
AgeValue = StringVar()

nameEntry = Entry(root, textvariable=nameValue, width=45, bd=2, font=20)
contactEntry = Entry(root, textvariable=contactValue, width=45, bd=2, font=20)
ageEntry = Entry(root, textvariable=AgeValue, width=15, bd=2, font=20)

# gender
gender_combobox = Combobox(root,
                           values=['Male', 'Female', 'Others'],
                           font='arial 14',
                           state='r',
                           width=14)
gender_combobox.place(x=450, y=200)
gender_combobox.set('Select')

addressEntry = Text(root, width=50, height=4, bd=2)

nameEntry.place(x=160, y=100)
contactEntry.place(x=160, y=150)
ageEntry.place(x=160, y=200)
addressEntry.place(x=160, y=250)

Button(root,
       text="Submit",
       bg="#326373",
       fg="white",
       width=13,
       height=1,
       command=submit).place(x=200, y=350)
Button(root,
       text="Clear",
       bg="#326373",
       fg="white",
       width=13,
       height=1,
       command=clear).place(x=340, y=350)
Button(root,
       text="Exit",
       bg="#326373",
       fg="white",
       width=13,
       height=1,
       command=lambda: root.destroy()).place(x=480, y=350)

root.mainloop()
