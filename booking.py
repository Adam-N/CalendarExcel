from tkinter import *
import datetime
from tkinter import messagebox

import xlsxwriter
from tkcalendar import Calendar


class Window(Frame):

    def __init__(self, master=None):
        Frame.__init__(self, master)

        # initializing for input how many periods in a day
        self.period_input = Spinbox(from_=3, to=8, wrap=True)

        # initializing for putting the start day for the calendar.
        self.start_date_input = Calendar(font='Arial 10', showweeknumbers=False)
        self.skip_days_input = Calendar(font='Arial 10', showweeknumbers=False)
        self.skip_days_input.bind("<<CalendarSelected>>", self.date_add)

        self.skip_days = []

        # variable for the start date
        self.start_date = datetime.datetime

        self.teach_names = ""
        self.teach_input = Entry(textvariable=self.teach_names, width=20)

        # to create input for the title of the spreadsheet.
        self.title = ""
        self.title_input = Entry(textvariable=self.title, width=20, )
        self.title_input_label = Label(text='Name for the file.')

        # for day cycle amount
        self.day_cycle = Spinbox(from_=1, to=8, wrap=True)
        self.day_cycle_window = 1

        self.schedule_day_full = {}
        self.repeating_window_cycle = 1

        week_counter_default = StringVar()
        week_counter_default.set("44")
        self.week_counter = Spinbox(from_=1, to=52, wrap=True, textvariable=week_counter_default)
        self.week_counter_label = Label(text='Number of weeks:')

        self.master = master
        self.init_window()

    # Creation of window
    def init_window(self):
        # changing the title of our master widget
        self.master.title("Computer Lab Booking Generator")

        self.pack(fill=BOTH, expand=1)

        # creating a button to quit the program
        quit_button = Button(self, text="Quit", command=self.client_exit)
        quit_button.place(x=9, y=590)

        # To input the amount of periods in a school day.
        self.period_input.place(x=150, y=50, width=30)
        period_input_label = Label(text='Number of Periods')
        period_input_label.place(x=1, y=50)

        # button to run the generate the spreadsheet.
        generate_button = Button(self, text='Generate', command=self.open_window)
        generate_button.place(x=250, y=590)

        # placing the calendar to pick start date.
        self.start_date_input.place(x=25, y=150)
        self.skip_days_input.place(x=25, y=375)
        start_date_label = Label(text='Please select a day to start your schedule on.')
        start_date_label.place(x=1, y=125)

        # placing the input for teacher names and it's label.
        self.teach_input.place(x=150, y=25)
        teach_input_label = Label(text='Teachers Names: ')
        teach_input_label.place(x=1, y=25)

        skip_days_input_label = Label(
            text=' Select days that there is no school(ie PD Days, Holidays, etc) ', justify='left')
        skip_days_input_label.place(x=1, y=350)

        self.title_input.place(x=150, y=1)
        self.title_input_label.place(x=1, y=1)
        self.title_input.insert(END, 'Lab Booking')

        self.day_cycle.place(x=150, y=75, width=30)
        day_cycle_label = Label(text='Number of days in a cycle:')
        day_cycle_label.place(x=1, y=73)

        # Check box. Will show the day on each date if checked.
        self.show_day_number_check = IntVar()
        show_day_number = Checkbutton(root, text="Show day number on each day",
                                      variable=self.show_day_number_check, onvalue=1, offvalue=0,
                                      height=1, width=25)
        show_day_number.place(x=1, y=557)

        # Spinbox for selecting how many weeks the program will run for
        self.week_counter.place(x=150, y=100, width=30)
        self.week_counter_label.place(x=1, y=100)

    @staticmethod
    def client_exit():
        exit()

    # to convert from a datetime to excels epoch time.
    def excel_date(self):
        offset = 693594
        self.start_date = datetime.datetime.strptime(self.start_date_input.get_date(), '%Y-%m-%d')
        n = self.start_date.toordinal()
        return n - offset

    # excel epoch for skip list.
    def excel_date_skip_list(self):
        offset = 693594
        self.start_date = datetime.datetime.strptime(self.skip_days_input.get_date(), '%Y-%m-%d')
        n = self.start_date.toordinal()
        return n - offset

    # displays the dates beside the calendar used to pick dates with no school.
    def display_dates(self):
        date_list = []
        for dates in self.skip_days:
            date_list.append(self.regular_date(self, dates))
            dates_list_label = Label(text=date_list, wraplength=225, justify='center').place(x=275, y=425)

    # converts from an excel epoch time to datetime.
    @staticmethod
    def regular_date(self, date):
        offset = 693594
        new_date = date + offset
        final_date = datetime.datetime.strftime(datetime.datetime.fromordinal(new_date), '%b %d %Y')
        return final_date

    # makes a list of teachers names that was input into the GUI.
    def make_teach_list(self):
        self.teach_names = self.teach_input.get()
        teach_list = self.teach_names
        teach_list = [x.strip() for x in teach_list.split(',')]

    # adds days to a list in excel format to check against for formatting and day cycle counts.
    def date_add(self, date):
        added_date = self.excel_date_skip_list()
        if added_date in self.skip_days:
            self.skip_days.remove(added_date)
            self.display_dates()
            return
        if added_date not in self.skip_days:
            self.skip_days.append(added_date)
            self.display_dates()

    # Opens window for modifying each day individually.
    def open_window(self):
        self.repeating_window = Toplevel(root)
        repating_sched_label = Label(self.repeating_window,
                                     text='Repeating Schedule for Day {}'.format(self.day_cycle_window)).grid(row=1,
                                                                                                              column=1)
        self.p1_input = Entry(self.repeating_window)
        self.p2_input = Entry(self.repeating_window)
        self.p3_input = Entry(self.repeating_window)
        self.p4_input = Entry(self.repeating_window)
        self.p5_input = Entry(self.repeating_window)
        self.p6_input = Entry(self.repeating_window)
        self.p7_input = Entry(self.repeating_window)
        self.p8_input = Entry(self.repeating_window)

        self.repeating_window.geometry("300x200+300+300")
        placement_y = 2
        next_day_button = Button(self.repeating_window, text='Next Day', command=self.next_day).grid(row=19, column=2)

        for i in range(1, int(self.period_input.get()) + 1):
            period_label_new_window = Label(self.repeating_window, text='Period {}'.format(i)).grid(row=placement_y,
                                                                                                    column=1)
            if i == 1:
                self.p1_input.grid(row=placement_y, column=2)
                self.p1_input.insert(END, "Name:")

            if i == 2:
                self.p2_input.grid(row=placement_y, column=2)
                self.p2_input.insert(END, "Name:")

            if i == 3:
                self.p3_input.grid(row=placement_y, column=2)
                self.p3_input.insert(END, "Name:")

            if i == 4:
                self.p4_input.grid(row=placement_y, column=2)
                self.p4_input.insert(END, "Name:")

            if i == 5:
                self.p5_input.grid(row=placement_y, column=2)
                self.p5_input.insert(END, "Name:")

            if i == 6:
                self.p6_input.grid(row=placement_y, column=2)
                self.p6_input.insert(END, "Name:")

            if i == 7:
                self.p7_input.grid(row=placement_y, column=2)
                self.p7_input.insert(END, "Name:")

            if i == 8:
                self.p8_input.grid(row=placement_y, column=2)
                self.p8_input.insert(END, "Name:")

            placement_y += 2

    # Brings up the next day, and stores the values into a dict.
    def next_day(self):
        self.schedule_day = {1: self.p1_input.get(), 2: self.p2_input.get(), 3: self.p3_input.get(),
                             4: self.p4_input.get(), 5: self.p5_input.get(), 6: self.p6_input.get(),
                             7: self.p7_input.get(), 8: self.p8_input.get()}

        for i in range(1, len(self.schedule_day) + 1):
            if self.schedule_day.get(i) == 'Name:' or self.schedule_day.get(i) == '':
                self.schedule_day.update({i: 'default'})
            i += 1
        self.schedule_day_full[self.repeating_window_cycle] = self.schedule_day
        self.repeating_window_cycle += 1
        self.day_cycle_window += 1
        self.repeating_window.destroy()
        self.open_window()

        if self.day_cycle_window == int(self.day_cycle.get()) + 1:
            self.repeating_window.destroy()
            self.generate()
            self.day_cycle_window = 1
            self.schedule_day_full = {}
            self.repeating_window_cycle = 1
            return 'None'

    # generates the excel spreadsheet.
    def generate(self):
        # Create a workbook and add a worksheet.
        workbook = xlsxwriter.Workbook('{}.xlsx'.format(self.title_input.get()))
        link_page = workbook.add_worksheet('Link Page')
        worksheet = workbook.add_worksheet("Week1")
        self.start_date = datetime.datetime.strptime(self.start_date_input.get_date(), '%Y-%m-%d')
        self.make_teach_list()
        # Start from the first cell. Rows and columns are zero indexed.
        first_date_link_page = 0
        row = 1
        col = 1

        # Starts Weeks at num 1
        week = 1

        day = 1

        # getting the excel epoch time for the start date.
        start_date_excel = self.excel_date()
        # Cell formats
        date_format = workbook.add_format({'num_format': 'd mmm yyy'})
        center_format = workbook.add_format()
        colour_format = workbook.add_format()
        bold_format = workbook.add_format()
        center_format.set_center_across()
        colour_format.set_bg_color('red')
        bold_format.set_bold(True)

        # variables to format the period cells.
        period_row = 2
        period_number = int(self.period_input.get())

        link_page_row = 1
        link_page_column = 0
        link_page.write("A1", "Click These")
        link_page.set_column(0, 0, width=15)
        link_page.set_column(1, 1, width=30)

        # Checks to see if the start day falls between a monday and friday.
        if self.start_date.isoweekday() <= 5:

            # title for the week.
            worksheet.write(0, 0, 'Week {}'.format(week))

            # writes all the periods along the side at the top.
            for p in range(1, (period_number + 1)):
                worksheet.write(period_row, 0, 'Period {}'.format(p))
                period_row += 1

            for q in range(self.start_date.isoweekday(), 6):
                worksheet.write(row, col, start_date_excel, date_format)
                worksheet.set_column(col, col, 18, center_format)

                # checks if this is the first day of the week.
                if q == self.start_date.isoweekday():
                    first_date_link_page = start_date_excel

                # sets for putting the information in the periods.
                period_row = 3

                # formats to a red column if no school that day.
                if start_date_excel in self.skip_days:
                    worksheet.write(2, col, 'No School', colour_format)
                    for h in range(0, period_number - 1):
                        worksheet.write(period_row, col, " ", colour_format)
                        period_row += 1
                        h += 1

                # formats cells to use a list provided by user.
                if start_date_excel not in self.skip_days:
                    period = 1
                    for j in range(0, period_number - 1):
                        if int(self.day_cycle.get()) == 1 or day > int(self.day_cycle.get()):
                            day = 1
                        if self.schedule_day_full[day][period] == 'default':
                            if self.teach_names:
                                worksheet.data_validation(period_row - 1, col, period_row, col,
                                                          {'validate': 'list', 'source': [self.teach_names]})
                            period_row += 1
                            period += 1
                        else:
                            worksheet.write(period_row - 1, col, self.schedule_day_full[day][period])
                            period_row += 1
                            period += 1
                    if self.show_day_number_check.get() == 1:
                        worksheet.write(period_row, col, "Day {}".format(day))

                col += 1
                day += 1
                start_date_excel += 1
                q += 1

                # Adds internal link so that you can navigate from the first page of the document.
                if q == 5:
                    second_date_link_page = start_date_excel
                    link_page.write_url(link_page_row, link_page_column, "internal:Week{}!A1".format(week),
                                        center_format,
                                        'Week{}'.format(week))

                    link_page.write_rich_string(link_page_row, link_page_column + 1, date_format,
                                                '{}'.format(self.regular_date(self, first_date_link_page)), ' until ',
                                                date_format,
                                                '{}'.format(self.regular_date(self, second_date_link_page)))
                    link_page_row += 1

                # adds link to first page with all the links to pages.
                worksheet.write_url('A12', "internal:'Link Page'!A1", bold_format, "First Page")

            # creates new worksheet for the next week.
            week += 1
            worksheet = workbook.add_worksheet('Week{}'.format(week))

            # Adds two to account for the weekend.
            start_date_excel += 2

            # Resets variables for the next round. Because they need to be back in the same cells.
            row = 1
            col = 1
            period_row = 2

        #if the day picked is a saturday or sunday.
        elif self.start_date.isoweekday() == 6:
            start_date_excel += 2

        elif self.start_date.isoweekday() == 7:
            start_date_excel += 1

        # This is for all the following weeks.
        for r in range(1, int(self.week_counter.get())):
            # title for the week.
            worksheet.write(0, 0, 'Week {}'.format(week))

            # writes all the dates at the top.
            for p in range(1, (period_number + 1)):
                worksheet.write(period_row, 0, 'Period {}'.format(p))
                period_row += 1

            for i in range(0, 5):
                worksheet.write(row, col, start_date_excel, date_format)
                worksheet.set_column(col, col, 18, center_format)
                # checks if this is the first day of the week.
                if i == 0:
                    first_date_link_page = start_date_excel

                # sets for putting the information in the periods.
                period_row = 3

                # formats to a red column if no school that day.
                if start_date_excel in self.skip_days:
                    worksheet.write(2, col, 'No School', colour_format)
                    for h in range(0, period_number - 1):
                        worksheet.write(period_row, col, " ", colour_format)
                        period_row += 1
                        h += 1

                # formats cells to use a list provided by user.
                if start_date_excel not in self.skip_days:
                    period = 1
                    if int(self.day_cycle.get()) == 1 or day > int(self.day_cycle.get()):
                        day = 1

                    for j in range(0, period_number - 1):
                        if self.schedule_day_full[day][period] == 'default':
                            if self.teach_names:
                                worksheet.data_validation(period_row - 1, col, period_row, col,
                                                          {'validate': 'list', 'source': [self.teach_names]})
                            period_row += 1
                            period += 1
                        else:
                            worksheet.write(period_row - 1, col, self.schedule_day_full[day][period])
                            period_row += 1
                            period += 1
                    if self.show_day_number_check.get() == 1:
                        worksheet.write(period_row, col, "Day {}".format(day))

                # iterates variables
                col += 1
                start_date_excel += 1
                i += 1
                day += 1

                # Adds internal link so that you can navigate from the first page of the document.
                if i == 4:
                    second_date_link_page = start_date_excel
                    link_page.write_url(link_page_row, link_page_column, "internal:Week{}!A1".format(week),
                                        center_format,
                                        'Week{}'.format(week))

                    link_page.write_rich_string(link_page_row, link_page_column + 1, date_format,
                                                '{}'.format(self.regular_date(self, first_date_link_page)), ' until ',
                                                date_format,
                                                '{}'.format(self.regular_date(self, second_date_link_page)))
                    link_page_row += 1

            week += 1

            # adds link to first page with all the links to pages.
            worksheet.write_url('A12', "internal:'Link Page'!A1", bold_format, "First Page")

            # creates new worksheet for the next week.
            worksheet = workbook.add_worksheet('Week{}'.format(week))

            # Adds two to account for the weekend.
            start_date_excel += 2

            r += 1

            # Resets variables for the next round. Because they need to be back in the same cells.
            row = 1
            col = 1
            period_row = 2

        link_page.activate()
        try:
            workbook.close()

        except PermissionError:
            messagebox.showinfo(message='There has been an error. Try closing the excel spreadsheet and trying again.')
            return 'None'
        messagebox.showinfo(message='Completed. Thank you!')
        self.client_exit()


root = Tk()

# size of the window
root.geometry("500x625")
root.resizable(FALSE, FALSE)
app = Window(root)
root.mainloop()
