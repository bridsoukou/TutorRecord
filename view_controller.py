from tkinter import filedialog, ttk
import tkinter.messagebox as tkmb
from tkinter import *
import model as m
import calendar
import string
import time
import atexit
import pickle
# TODO: Add folder(s) for tr.ico, file_path.p, settings data, etc.
# TODO: Add a delete student button.
# TODO: Add profile ability so that people sharing the computer can have their own record files.
# TODO: Add export PayPal invoice (csv) function


class MainApplication:
    def __init__(self, master):
        atexit.register(self.on_closing)
        # master.protocol("WM_DELETE_WINDOW", root.iconify)
        # *** Setup Main Frame ***
        self.master = master

        main_frame = Frame(master)
        main_frame.grid(sticky="WENS")

        # Save record file path once selected
        try:
            self.file_name = pickle.load(open("file_path.p", "rb"))
        except FileNotFoundError:
            self.file_name = filedialog.askopenfilename(
                initialdir="/", title="Select record file",
                filetypes=(("Excel files", "*.xls *.xlsx"), ("all files", "*.*"))
            )
            self.file_name = self.file_name
            pickle.dump(self.file_name, open("file_path.p", "wb"))
        finally:
            self.model = m.Model(self.file_name)

        self.date_format = self.model.date_format

        # *** Setup Students Listbox ***
        self.table_frame = Frame(main_frame)
        self.table_frame.grid(row=1, padx=2, pady=2, sticky="WE")
        self.students_lb = Listbox(self.table_frame, selectmode=SINGLE, exportselection=False)
        self.students_lb.grid(row=0, column=0)
        self.populate_students()

        # *** Setup Info Panel ***
        self.dates_lb = Listbox(self.table_frame, width=10, exportselection=False)
        self.dates_lb.grid(row=0, column=1)
        self.remaining_label = Label(self.table_frame, text="Classes Remaining:")
        self.remaining_label.grid(row=0, column=2, sticky=N)
        self.booking_label = Label(self.table_frame, text="N/A", font=("courier", 32), relief=RIDGE)
        self.booking_label.grid(row=0, column=2)

        # Flip focus from Students_lb to dates_lb with left and right arrow keys.
        def bind_focus_set(f):
            f.focus_set()
        master.bind("<Right>", lambda f=self.dates_lb: bind_focus_set(self.dates_lb))
        master.bind("<Left>", lambda f=self.dates_lb: bind_focus_set(self.students_lb))

        # ***Setup Toolbar***
        toolbar = Frame(main_frame, bd=0)
        toolbar.grid(row=0, sticky=W)
        new_student_button = Button(toolbar, text="+Student", command=self.add_student_button)
        new_student_button.grid(row=0, padx=2, pady=2)
        choose_file_button = Button(toolbar, text="Choose Record File", command=self.open_filedialog)
        choose_file_button.grid(row=0, column=1, padx=2, pady=2)
        self.classes_on_day = ttk.Combobox(toolbar, values=list(range(1, 5)), width=1)
        self.classes_on_day.grid(row=0, column=3, padx=2, pady=2)

        self.had_class_var = IntVar()
        self.had_class_cb = Checkbutton(toolbar, text="Class done!", fg="green",
                                        command=self.mark_done, variable=self.had_class_var)
        self.had_class_cb.grid(row=0, column=2, padx=2, pady=2)

        # ***Setup Menu Bar ***
        menu = Menu(master)
        master.config(menu=menu)

        file_menu = Menu(menu, tearoff=0)
        file_menu.add_command(label="Select Record File", command=self.open_filedialog)
        file_menu.add_command(label="Refresh Record File", command=self.refresh_record)
        file_menu.add_command(label="Update Record File", command=self.flush_changes)

        def print_dates():
            student = self.model.students_list[self.get_student_from_name(self.students_lb.get(ACTIVE))]
            for date in student.dates:
                print(date)

        students_menu = Menu(menu, tearoff=0)
        students_menu.add_command(label="New Student", command=self.add_student_button)
        students_menu.add_command(label="Edit Student", command=self.edit_student_button)
        students_menu.add_separator()
        students_menu.add_command(label="Print Dates", command=print_dates)

        menu.add_cascade(label="File", menu=file_menu)
        menu.add_cascade(label="Students", menu=students_menu)

        self.populate_dates()
        self.update_ui()

    def add_student_button(self):
        AddOrEditStudent("add")
        try:
            self.populate_students()
        except TclError:
            pass

    def edit_student_button(self):
        AddOrEditStudent("edit", self.students_lb.get(ACTIVE))
        try:
            self.populate_students()
        except TclError:
            pass

    def mark_done(self):
        active_student_name = self.students_lb.get(ACTIVE)
        active_student_index = self.get_student_from_name(active_student_name)
        date = self.dates_lb.get(ACTIVE)
        struct_date = time.strptime(date, self.date_format)
        how_many_classes = int(self.classes_on_day.get())
        if self.had_class_var.get() == 0:
            self.dates_lb.itemconfig(active_student_index, bg="gray", fg="white")
            for x in range(how_many_classes):
                self.model.students_list[active_student_index].dates.remove(struct_date)
            print("Date removed.")
            self.model.students_list[active_student_index].booking += how_many_classes
        else:
            self.dates_lb.itemconfig(active_student_index, bg="#59f766", fg="white")
            for x in range(how_many_classes):
                self.model.students_list[active_student_index].dates.append(struct_date)
            self.model.students_list[active_student_index].booking -= how_many_classes
            print("Date added.")
        self.populate_booking(self.model.students_list[active_student_index])

    def get_student_from_name(self, name):
        for x, student in enumerate(self.model.students_list):
            if student.name == name:
                return x

    def populate_booking(self, student):
        try:
            value = int(student.booking)
        except ValueError:
            value = "N/A"

        if value < 1:
            self.booking_label.config(text=value, bg="red", bd=5)
        elif value < 5:
            self.booking_label.config(text=value, bg="yellow", bd=5)
        else:
            self.booking_label.config(text=value, bg="green", bd=5)

        self.booking_label.config(text=value)

    def populate_dates(self):
        name = self.students_lb.get(ACTIVE)
        for student in self.model.students_list:
            if student.name == name:
                self.populate_booking(student)
                self.dates_lb.delete(0, END)
                # Put the newest date at the top, then put them in the Listbox
                student.class_dates_until_today = self.model.calculate_class_days(student)
                dates = list(reversed(student.class_dates_until_today))
                for date in dates:
                    self.dates_lb.insert(END, time.strftime(self.date_format, date))

    def populate_students(self):
        if self.students_lb.size() > 0:
            self.students_lb.delete(0, END)
        for student in self.model.students_list:
            self.students_lb.insert(END, student.name)

    def open_filedialog(self):
        file_path = filedialog.askopenfilename(
            initialdir="/", title="Select record file",
            filetypes=(("Excel files", "*.xls *.xlsx"), ("all files", "*.*"))
        )
        self.refresh_record()
        self.populate_students()
        return file_path

    def refresh_record(self):
        if self.file_name is not "":
            self.model = m.Model(self.file_name)

    def update_ui(self):
        try:
            if self.master.focus_get() == self.classes_on_day:
                pass
            elif self.master.focus_get() == self.students_lb:
                self.populate_dates()
                self.students_lb.select_clear(0, END)
                self.students_lb.select_set(self.students_lb.index(ACTIVE))
            elif self.master.focus_get() == self.dates_lb:
                self.dates_lb.select_clear(0, END)
                self.dates_lb.select_set(self.dates_lb.index(ACTIVE))
                active_student_index = self.get_student_from_name(self.students_lb.get(ACTIVE))

                # Make the had_class checkbox reflect whether the student had class on this date or not.
                active_date_struct = time.strptime(self.dates_lb.get(ACTIVE), self.date_format)
                current_student_dates = self.model.students_list[active_student_index].dates
                if active_date_struct in current_student_dates:
                    self.had_class_var.set(True)
                else:
                    self.had_class_var.set(False)

                # Get the number of classes student is supposed to have on this day and put it in the
                # classes_on_day combobox.
                self.classes_on_day.set(self.count_classes_on_this_day())
        except KeyError:
            # Error while combo box drop down is open. No idea why.
            print(KeyError)

        # Mark dates green if they exist in the record, white if not
        index = self.students_lb.index(ACTIVE)
        struct_date_items_dates_lb = [time.strptime(i, self.date_format) for i in self.dates_lb.get(0, END)]
        for i, item in enumerate(struct_date_items_dates_lb):
            if item in self.model.students_list[index].dates:
                # Turn them light green with white text
                self.dates_lb.itemconfig(i, bg="#59f766", fg="white")
            else:
                # Turn them gray with white text
                self.dates_lb.itemconfig(i, bg="gray", fg="white")

        self.master.after(100, self.update_ui)

    def count_classes_on_this_day(self):
        # Get the number of classes student is supposed to have on this day
        active_student_name = self.students_lb.get(ACTIVE)
        active_student_index = self.get_student_from_name(active_student_name)
        active_date_struct = time.strptime(self.dates_lb.get(ACTIVE), self.date_format)
        active_student_cpd = self.model.students_list[active_student_index].classes_per_day
        classes_on_this_day = active_student_cpd[active_date_struct[6]]
        return classes_on_this_day

    def flush_changes(self):
        try:
            self.model.update_sheet()
            self.model.save_sheet()
        except PermissionError:
            tkmb.showerror("Error", "Please close the record file and try again!")
            self.model.save_sheet()

    def on_closing(self):
        self.flush_changes()


class AddOrEditStudent:
    def __init__(self, action, editee=None):
        self.window = Tk()
        self.window.title("Add or Edit Student")
        self.window.iconbitmap("tr.ico")
        self.frame = Frame(self.window)
        self.frame.grid(padx=2, pady=2)
        self.editee_index = None
        self.action = action
        self.editee = editee

        self.model = application.model

        self.name_label = Label(self.frame, text="Name")
        self.name_label.grid(row=0, padx=2, pady=2)

        self.name_entry = Entry(self.frame)
        self.name_entry.grid(row=0, column=1, columnspan=2, padx=2, pady=2)

        if action == "edit":
            for student in self.model.students_list:
                if student.name == editee:
                    self.editee_index = self.model.students_list.index(student)
                    self.name_entry.insert(0, editee)

        # Setup a combo box for each day of the week.
        self.day_cboxes = []
        for d in range(0, 7):
            day = Label(self.frame, text=calendar.day_abbr[d])
            day.grid(row=d + 2)
            self.cbox = ttk.Combobox(self.frame, width=2, values=list(range(0, 5)))
            self.day_cboxes.append(self.cbox)
            self.cbox.grid(row=d + 2, column=1, sticky=W, padx=2, pady=2, ipadx=1, ipady=1)
            if action == "add":
                self.cbox.set(0)
            else:
                self.cbox.set(self.model.students_list[self.editee_index].classes_per_day[d])

        booking_label = Label(self.frame, text="Booking")
        booking_label.grid(column=2, row=2, sticky=W)

        self.booking_cbox = ttk.Combobox(self.frame, width=2, values=(10, 25, 50))
        self.booking_cbox.grid(column=2, row=3, sticky=W, padx=2, pady=2)
        if action == "add":
            self.booking_cbox.set(20)
        else:
            self.booking_cbox.set(self.model.students_list[self.editee_index].booking)

        save_button = Button(self.frame, text="Save", command=self.save_button)
        save_button.grid(columnspan=2, padx=2, pady=2, ipadx=1, ipady=1)
        cancel_button = Button(self.frame, text="Cancel", command=self.my_quit)
        cancel_button.grid(columnspan=2, padx=2, pady=2, ipadx=1, ipady=1)

        self.window.mainloop()

    def my_quit(self):
        self.window.quit()
        self.window.destroy()

    def save_button(self):
        if self.action == "add":
            new_student = m.Student()
            new_student.name = self.name_entry.get()
            new_student.booking = int(self.booking_cbox.get())
            new_student.dates = []
            # TODO: Rework this using row and column notation, where column can be an int instead of a char
            try:
                new_student.column = string.ascii_lowercase[len(self.model.students_list) + 1]
            except IndexError:
                new_student.column = "A" + string.ascii_lowercase[1]
            new_student.classes_per_day = [i.get() for i in self.day_cboxes]

            self.model.students_list.append(new_student)
        else:
            index = self.editee_index
            self.model.students_list[index].name = self.name_entry.get()
            self.model.students_list[index].booking = self.booking_cbox.get()
            self.model.students_list[index].classes_per_day = [i.get() for i in self.day_cboxes]

        self.my_quit()


root = Tk()

root.title("TutorRecord")
root.iconbitmap("tr.ico")
root.resizable(False, False)
application = MainApplication(root)

root.mainloop()
