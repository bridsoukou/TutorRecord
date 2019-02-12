import openpyxl
import re
import time


class Student:
    def __init__(self):
        self.name = None
        self.booking = None
        self.dates = []
        self.column = None    # M, T, W, T, F, S, S
        self.classes_per_day = [0, 0, 0, 0, 0, 0, 0]
        self.class_dates_until_today = []


class Model:
    def __init__(self, file_path):
        self.file_name = file_path

        self.workbook = openpyxl.load_workbook(file_path)
        self.sheet = self.workbook.active

        self.col_row_regex = re.compile(r"<Cell '\S+'.([A-Z])(\d+)>")
        self.date_format = "%m/%d/%Y"
        self.alt_date_format = "%Y-%m-%d"
        self.name_row = int("1")
        self.booking_row = int("2")
        self.schedcode_row = int("3")
        self.dates_start_row = int("4")
        self.ignore_column = "A"
        self.students_list = []

        self.process_sheet()
        self.print_students()

    def process_sheet(self):
        self.read_sheet()
        for student in self.students_list:
            student.class_dates_until_today = self.calculate_class_days(student)

    def read_sheet(self):
        if len(self.students_list) > 1:
            self.students_list.clear()
        # Regex to scan the column and put the column in group 1, the row in group 2.
        cols = self.sheet.columns

        for col in cols:
            # If the column is set to be skipped because it contains labels, skip it!
            mo = self.col_row_regex.search(str(col))
            current_column = mo.group(1)
            if current_column == self.ignore_column:
                continue

            new_student = Student()
            # Assign the new student a column if he doesn't have one.
            if new_student.column is None:
                new_student.column = mo.group(1)

            for cell in col:
                if cell.value is not None:
                    mo = self.col_row_regex.search(str(cell))
                    current_row = int(mo.group(2))
                    if current_row == self.name_row:
                        new_student.name = cell.value
                    elif current_row == self.booking_row:
                        new_student.booking = cell.value
                    elif current_row == self.schedcode_row:
                        new_student.classes_per_day = [int(i) for i in cell.value if i != "s"]
                    elif current_row >= self.dates_start_row:
                        year_month_day = str(cell.value)[:10]
                        try:
                            struct_time_date = time.strptime(year_month_day, self.date_format)
                        except ValueError:
                            struct_time_date = time.strptime(year_month_day, self.alt_date_format)
                            print("Alt format used")
                        new_student.dates.append(struct_time_date)

            if new_student.name is not None:
                self.students_list.append(new_student)

    def update_sheet(self):
        for student in self.students_list:
            # Update name
            name_cell = self.sheet[str(student.column + str(self.name_row))]
            name_cell.value = student.name
            # Update booking
            booking_cell = self.sheet[student.column + str(self.booking_row)]
            booking_cell.value = student.booking
            # Update schedcode
            schedcode_cell = self.sheet[student.column + str(self.schedcode_row)]
            for i, cpd in enumerate(student.classes_per_day):
                student.classes_per_day[i] = str(cpd)
            schedcode = "s" + "".join(student.classes_per_day)
            schedcode_cell.value = schedcode
            # Update dates
            # Clear old dates
            for x, row in enumerate(self.sheet.rows):
                date_cell = self.sheet[student.column + str(self.dates_start_row + x)]
                date_cell.value = None
            # Add new dates
            for x, date in enumerate(student.dates):
                date_cell = self.sheet[student.column + str(self.dates_start_row + x)]
                date_cell.value = time.strftime(self.date_format, date)
        print("Sheet updated.")

    def save_sheet(self):
        self.workbook.save(self.file_name)
        print("Workbook saved")

    def print_students(self):
        for student in self.students_list:
            print(student.name)
            print(student.booking)
            for date in student.dates:
                print(time.strftime(self.date_format, date))

    @staticmethod
    def calculate_class_days(student):
        class_dates_until_today = []
        if student.classes_per_day is not []:
            year_day_format = "Year: %Y Day: %j"
            now = time.localtime()
            year = now[0]
            today = now[7] + 1
            classes_on_day = student.classes_per_day
            # For every day of the year until today, parse the date.
            for xday in range(1, today):
                day_str = "Year: %s Day: %s" % (year, xday)
                day = time.strptime(day_str, year_day_format)
                week_day = day[6]
                # If the student has classes on that date, append the date to the list of dates until today.
                if classes_on_day[week_day] != 0 and classes_on_day[week_day] != "0":
                    class_dates_until_today.append(day)

        return class_dates_until_today
