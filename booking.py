from tkinter import *
import xlsxwriter
from tkcalendar import Calendar, DateEntry
import datetime


class Window(Frame):

    def __init__(self, master=None):
        Frame.__init__(self, master)

        # initializing for input how many periods in a day
        self.period_input = Spinbox(from_=3, to=8, wrap=True)

        # Day cycle input. How many days does the school cycle around High typically 2, elem typically 6.
        # Unused ATM - Future implementation
        # self.day_cycle_input = Spinbox(from_=1, to=8, wrap=True)

        # initializing for putting the start day for the calendar.
        self.start_date_input = Calendar(font='Arial 10', showweeknumbers=False)
        self.skip_days_input = Calendar(font='Arial 10', showweeknumbers=False)
        self.skip_days_input.place(x=25, y=350)
        self.skip_days_input.bind("<<CalendarSelected>>", self.date_add)

        self.skip_days = []

        # variable for the start date
        self.start_date = datetime.datetime

        # Label for an alert if the first day is not a monday
        self.no_monday_label = Label(text='Please choose Monday as your starting day.')
        self.teach_names = ""
        self.teach_input = Entry(textvariable=self.teach_names, width=20)

        # to create input for the title of the spreadsheet.
        self.title = ""
        self.title_input = Entry(textvariable=self.title, width=20, )
        self.title_input_label = Label(text='Name for the file.')

        # testing start date entry
        self.master = master
        self.init_window()

    # Creation of init_window
    def init_window(self):
        # changing the title of our master widget
        self.master.title("Computer Lab Booking Generator")

        self.pack(fill=BOTH, expand=1)

        # creating a button to quit the program
        quit_button = Button(self, text="Quit", command=self.client_exit)
        quit_button.place(x=9, y=570)

        # To input the amount of periods in a school day.
        self.period_input.place(x=150, y=50, width=30)
        period_input_label = Label(text='Number of Periods')
        period_input_label.place(x=1, y=50)

        # button to run the generate the spreadsheet.
        generate_button = Button(self, text='Generate', command=self.generate)
        generate_button.place(x=250, y=570)

        # placing the calendar to pick start date.
        self.start_date_input.place(x=25, y=125)
        start_date_label = Label(text='Please select a day to start your schedule one. Must be a monday.')
        start_date_label.place(x=1, y=100)

        # placing the input for teacher names and it's label.
        self.teach_input.place(x=150, y=1)
        teach_input_label = Label(text='Teachers Names: ')
        teach_input_label.place(x=1, y=1)

        skip_days_input_label = Label(text='Select days that there is no school(ie PD Days, Holidays, etc)')
        skip_days_input_label.place(x=1, y=325)

        # working on placing these
        # self.title_input.place(x=1, y=350)
        # self.title_input_label.place(x=1, y=340)

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

    # converts from an excel epoch time to datetime.
    def regular_date(self, date):
        offset = 693594
        new_date = date + offset
        final_date = datetime.datetime.strftime(datetime.datetime.fromordinal(new_date), '%b %d %Y')
        return final_date

    # checks to see if the date you picked was a monday
    def monday_check(self):
        self.start_date = datetime.datetime.strptime(self.start_date_input.get_date(), '%Y-%m-%d')
        if self.start_date.isoweekday() == 1:
            return True
        elif self.start_date.isoweekday() != 1:
            return False

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
            return
        if added_date not in self.skip_days:
            self.skip_days.append(added_date)

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

        # getting the excel epoch time for the start date.
        start_date_excel = self.excel_date()
        # Cell formats
        date_format = workbook.add_format({'num_format': 'd mmm yyy'})
        centre_format = workbook.add_format()
        colour_format = workbook.add_format()
        centre_format.set_center_across()
        colour_format.set_bg_color('red')

        # variables to format the period cells.
        period_row = 2
        period_number = int(self.period_input.get())

        link_page_row = 1
        link_page_column = 0
        link_page.write("A1", "Click These")

        if not self.monday_check():
            self.no_monday_label.place(x=50, y=325)
            return "None"

        if self.monday_check():
            self.no_monday_label.grid_forget()

        for r in range(1, 44):
            # title for the week.
            worksheet.write(0, 0, 'Week {}'.format(week))

            # writes all the dates at the top.
            for p in range(1, (period_number + 1)):
                worksheet.write(period_row, 0, 'Period {}'.format(p))
                period_row += 1

            for i in range(0, 5):
                worksheet.write(row, col, start_date_excel, date_format)

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
                    worksheet.data_validation(period_row, col, period_row + period_number, col + 1,
                                              {'validate': 'list', 'source': [self.teach_names]})

                # iterates variables
                col += 1
                start_date_excel += 1
                i += 1

                # Adds internal link so that you can navigate from the first page of the document.
                if i == 4:
                    second_date_link_page = start_date_excel
                    link_page.write_url(link_page_row, link_page_column, "internal:Week{}!A1".format(week),
                                        centre_format,
                                        'Week{}'.format(week))

                    link_page.write_rich_string(link_page_row, link_page_column + 1, date_format,
                                                '{}'.format(self.regular_date(first_date_link_page)), ' until ',
                                                date_format,
                                                '{}'.format(self.regular_date(second_date_link_page)))
                    link_page_row += 1

            week += 1

            # adds link to first page with all the links to pages.
            worksheet.write_url('A12', "internal:'Link Page'!A1", centre_format, "First Page")

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
        workbook.close()


root = Tk()

# size of the window
root.geometry("400x600")

app = Window(root)
root.mainloop()
