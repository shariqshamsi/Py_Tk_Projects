import tkinter
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl
import sqlite3

def enter_data():
    accepted = accept_var.get()
    if accepted == "Accepted":
        #user info
        firstname = first_name_entry.get()
        lastname = last_name_entry.get()
        
        if firstname and lastname:
            title = title_combobox.get()
            age = age_spinbox.get()
            nationality = nationality_combobox.get()

            #course info
            registration_status = reg_status_var.get()
            numcourses = numcourses_spinbox.get()
            numsemesters = numsemester_spinbox.get()

            print("First name : ", firstname, "Last Name : ", lastname)
            print("Title : ", title, "Age : ", age, "Nationality : ", nationality)
            print("# Courses: ", numcourses, "# Semester : ", numsemesters)
            print("Registration Staus : ", registration_status)
            print("----------------------------------------------------------------")
        
            conn = sqlite3.connect('data.db')
            table_create_query = '''CREATE TABLE IF NOT EXISTS Student_Data
                    (firstname TEXT, lastname TEXT, title TEXT, age INT, nationality TEXT,
                    registration_status TEXT, num_courses INT, num_semesters INT)
            '''
            conn.execute(table_create_query)
            
            #Insert Data
            data_insert_query = '''INSERT INTO Student_Data (firstname, 
            lastname,title,age,nationality,registration_status,num_courses,
            num_semesters)VALUES(?,?,?,?,?,?,?,?)'''
            
            data_insert_tuple = (firstname,lastname,title,age,nationality,
                                 registration_status,numcourses,numsemesters)
            
            cursor = conn.cursor()
            cursor.execute(data_insert_query,data_insert_tuple)
            conn.commit()
            
            
            
            
            conn.close()
            
            filepath = "H:\my project\Tkinter_projects\Tkinter_data_entry_form\data.xlsx"
            
            if not os.path.exists(filepath):
               workbook = openpyxl.Workbook() 
               sheet = workbook.active
               heading = ["First Name","Last Name","Title","Age","Nationality",
                          "#Course","#Semester","Registration Status"]
               sheet.append(heading)
               workbook.save(filepath)
            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook.active
            sheet.append([firstname, lastname, title,age, nationality,
                          numcourses,numsemesters,registration_status])
            workbook.save(filepath)
        
        
        
        else:
             tkinter.messagebox.showwarning(title = "Error", message = "First & last name are required")   
    else:
        tkinter.messagebox.showwarning(title= "Error", message= "You have not accepted terms & conditions")
    
    
window = tkinter.Tk()
window.title("Data Entry Form")

frame = tkinter.Frame(window)
frame.pack()

#saving user info
user_info_frame = tkinter.LabelFrame(frame, text = "User Information")
user_info_frame.grid(row = 0, column=0, padx= 20, pady=10)

first_name_label = tkinter.Label(user_info_frame, text = "First Name")
first_name_label.grid(row=0 , column=0)
last_name_label = tkinter.Label(user_info_frame, text = "Last Name")
last_name_label.grid(row=0 , column=1)

first_name_entry = tkinter.Entry(user_info_frame)
last_name_entry = tkinter.Entry(user_info_frame)
first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)

title_label = tkinter.Label(user_info_frame, text="Title")
title_combobox = ttk.Combobox(user_info_frame, values =["Mr.", "Ms."] )
title_label.grid(row=0, column=2)
title_combobox.grid(row=1, column=2)

age_label = tkinter.Label(user_info_frame, text = "Age")
age_spinbox = tkinter.Spinbox(user_info_frame, from_=18 , to = 110)
age_label.grid(row= 2, column=0)
age_spinbox.grid(row=3, column=0)

nationality_label = tkinter.Label(user_info_frame, text = "Nationality" )
nationality_combobox = ttk.Combobox(user_info_frame, values = ["Karachi", "Lahore","Multan"])
nationality_label.grid(row=2, column=1)
nationality_combobox.grid(row=3, column=1)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx = 10, pady = 5)


#saving course info
courses_frame = tkinter.LabelFrame(frame)
courses_frame.grid(row= 1, column=0, sticky="news", padx=20, pady=10)

registered_label = tkinter.Label(courses_frame, text="Registration Status")

reg_status_var = tkinter.StringVar(value = "Not Registered")
registered_check = tkinter.Checkbutton(courses_frame, text="Currently Registered", variable=reg_status_var, onvalue= "Registered",offvalue="Not Registered"  )

registered_label.grid(row=0, column=0)
registered_check.grid(row=1, column=0)

numcourses_label = tkinter.Label(courses_frame, text = "#completed courses")
numcourses_spinbox =tkinter.Spinbox(courses_frame, from_=0 ,to='infinity')
numcourses_label.grid(row=0, column=1)
numcourses_spinbox.grid(row=1, column=1)

numsemesters_label = tkinter.Label(courses_frame, text="#Semester")
numsemester_spinbox =tkinter.Spinbox(courses_frame, from_=0 ,to='infinity')
numsemesters_label.grid(row=0, column=2)
numsemester_spinbox.grid(row=1, column=2)

for widget in courses_frame.winfo_children():
    widget.grid_configure(padx = 10, pady = 5)
    
# Accept Terms
terms_frame = tkinter.LabelFrame(frame, text = "Terms & Conditions")
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

accept_var = tkinter.StringVar(value = "Not accepted")
terms_check= tkinter.Checkbutton(terms_frame, text="I accept all", variable = accept_var, onvalue= "Accepted", offvalue= "Not Accepted")
terms_check.grid(row=0, column=0)

#button
button = tkinter.Button(frame, text="Enter Data", command= enter_data)
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)


window.mainloop()