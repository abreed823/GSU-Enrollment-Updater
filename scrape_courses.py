import pandas as pd
from openpyxl import load_workbook
from datetime import date

# -TO DO-
# Distinguish multicast from embedded
# Figure out cell width
# Center text in cells
# Add proper formatting (i.e. cell outlines, titles, legends, etc.)
# Get proper page spacing


class Courses:
    # Creates new sheet
    today = date.today().strftime('%m-%d-%Y')
    book = load_workbook('Fall Schedule March 25th copy.xlsx')
    book.create_sheet('NEW TEST SHEET ' + today)
    book.save('Fall Schedule March 25th copy.xlsx')
    new_sheet = book.worksheets[-1]

    # Sets sheet column widths
    new_sheet.column_dimensions['A'].width = 5
    new_sheet.column_dimensions['B'].width = 6
    new_sheet.column_dimensions['C'].width = 5.17
    new_sheet.column_dimensions['D'].width = 9.67
    new_sheet.column_dimensions['E'].width = 8.67
    new_sheet.column_dimensions['F'].width = 29.17
    new_sheet.column_dimensions['G'].width = 4.67
    new_sheet.column_dimensions['H'].width = 17.67
    new_sheet.column_dimensions['I'].width = 4.67
    new_sheet.column_dimensions['J'].width = 5.67
    new_sheet.column_dimensions['K'].width = 9
    new_sheet.column_dimensions['L'].width = 16.17

    # Declares starting rows for each campus
    alp_row = 4
    clk_row = 34
    dec_row = 64
    dun_row = 94
    newt_row = 124
    onl_row = 154

    # Defines column of each piece of data in data frame
    col_crn = 2
    col_subj = 3
    col_class = 4
    col_sec = 5
    col_campus = 6
    col_credits = 7
    col_title = 8
    col_days = 9
    col_time = 10
    col_cap = 11
    col_act = 12
    col_prof = 18
    col_location = 20

    def row_has_class(self, crn):
        return any(char.isdigit() for char in crn)

    def add_data(self, crn, subj, class_name, sec, class_credits, title, days, class_time, cap, act, prof,
                 location, row_type):
        self.new_sheet.cell(row=row_type, column=2, value=crn)
        self.new_sheet.cell(row=row_type, column=3, value=subj)
        if cap < 19:
            self.new_sheet.cell(row=row_type, column=4, value=f'{class_name}-{sec}*')
        else:
            self.new_sheet.cell(row=row_type, column=4, value=f'{class_name}-{sec}')
        self.new_sheet.cell(row=row_type, column=5, value=class_credits)
        self.new_sheet.cell(row=row_type, column=6, value=title.upper())
        self.new_sheet.cell(row=row_type, column=7, value=days)
        self.new_sheet.cell(row=row_type, column=8, value=class_time.upper())
        if act != 0:
            self.new_sheet.cell(row=row_type, column=9, value=act)
        self.new_sheet.cell(row=row_type, column=10, value=cap)
        self.new_sheet.cell(row=row_type, column=11, value=location)

        if prof == 'TBA':
            self.new_sheet.cell(row=row_type, column=12, value='STAFF')
        else:
            original_name = prof.split(' ')
            first_initial = original_name[0][0]
            last_name = original_name[-1]
            # Maybe try this with list comprehension???
            if '-' in last_name:
                # Maybe try combining these two statements???
                hyphenated_name = last_name.split('-')
                last_name = hyphenated_name[-1]
            formatted_name = f'{first_initial}. {last_name}'
            self.new_sheet.cell(row=row_type, column=12, value=formatted_name[0:-4].upper())

        # The faster less brute-force version if I can get it to work
        # I need it to be able to add the list of values to specific rows, not just to the end of the sheet
        # new_row = [crn, subj, class_name, sec, campus, class_credits, title, days, class_time, cap, act, prof, location]
        # self.new_sheet[self.row].value = new_row
        # self.row += 1
        # print('Row constructed successfully!')

        # Testing purposes only
        # print(f'{crn} {subj} {class_name} {sec} {campus} {class_credits} {title} {days} {class_time} {cap} {act} {prof} {location}')

    def create_data_frame(self, html_source):
        # Pulls second to last table from site
        # 'header=0' allows my to properly label each column
        df = pd.read_html(html_source, header=0)[-2]

        for i in range(0, df.shape[0]):
            if self.row_has_class(df.iat[i, self.col_crn]):
                crn = df.iat[i, self.col_crn]
                subj = df.iat[i, self.col_subj]
                class_name = df.iat[i, self.col_class]
                sec = df.iat[i, self.col_sec]
                campus = df.iat[i, self.col_campus]
                class_credits = df.iat[i, self.col_credits]
                title = df.iat[i, self.col_title]
                days = df.iat[i, self.col_days]
                class_time = df.iat[i, self.col_time]
                cap = df.iat[i, self.col_cap]
                act = df.iat[i, self.col_act]
                prof = df.iat[i, self.col_prof]
                location = df.iat[i, self.col_location]

                # room_number = [i for i in location.split() if i.isdigit()][0]

                if 'Alpharetta' in campus:
                    self.alp_row += 1
                    row_type = self.alp_row
                    room_number = [i for i in location.split() if i.isdigit()][0]

                    if 'Bldg A' in location:
                        location = f'AA {room_number}'
                    else:
                        location = f'AB {room_number}'
                elif 'Clarkston' in campus:
                    self.clk_row += 1
                    row_type = self.clk_row
                    room_number = [i for i in location.split() if i.isdigit()][0]

                    if 'Bldg B' in location:
                        location = f'CB {room_number}'
                    elif 'Bldg C' in location:
                        location = f'CC {room_number}'
                    elif 'Bldg D' in location:
                        location = f'CD {room_number}'
                    elif 'Bldg E' in location:
                        location = f'CE {room_number}'
                    else:
                        location = f'CH {room_number}'
                elif 'Decatur' in campus:
                    self.dec_row += 1
                    row_type = self.dec_row
                    room_number = [i for i in location.split() if i.isdigit()][0]

                    if 'Bldg. SB' in location:
                        location = f'SB {room_number}'
                    else:
                        location = f'SC {room_number}'
                elif 'Dunwoody' in campus:
                    self.dun_row += 1
                    row_type = self.dun_row
                    room_number = [i for i in location.split() if i.isdigit()][0]

                    if 'Classroom' in location:
                        location = f'E {room_number}'
                    elif 'Science' in location:
                        location = f'C {room_number}'
                    else:
                        location = f'A {room_number}'
                elif 'Newton' in campus:
                    self.newt_row += 1
                    row_type = self.newt_row
                    room_number = [i for i in location.split() if i.isdigit()][0]
                    location = f'1N {room_number}'
                else:
                    self.onl_row += 1
                    row_type = self.onl_row

                self.add_data(int(crn), subj, class_name, sec, int(float(class_credits)), title, days, class_time,
                              int(cap), int(act), prof,
                              location, row_type)
            else:
                print(f'I am not a row. Index: {i}')
        self.book.save('Fall Schedule March 25th copy.xlsx')

        # ***** The slow version of the above code *****
        # for i in df.index:
        #     if self.row_has_class(df[col_crn].iloc[i]):
        #         # print('I am a row')
        #         crn = df[col_crn].iloc[i]
        #         subj = df[col_subj].iloc[i]
        #         class_name = df[col_class].iloc[i]
        #         sec = df[col_sec].iloc[i]
        #         campus = df[col_campus].iloc[i]
        #         class_credits = df[col_credits].iloc[i]
        #         title = df[col_title].iloc[i]
        #         days = df[col_days].iloc[i]
        #         class_time = df[col_time].iloc[i]
        #         cap = df[col_cap].iloc[i]
        #         act = df[col_act].iloc[i]
        #         prof = df[col_prof].iloc[i]
        #         location = df[col_location].iloc[i]
        #         self.add_data(crn, subj, class_name, sec, campus, class_credits, title, days, class_time, cap, act,
        #                       prof, location)
        #     else:
        #         print(f'I am not a row. Index: {i}')
        #     if i == 15:
        #         break

        # *****NOTE - DO NOT DELETE***** This block of code puts data frame into sheet.
        # This does not need to happen every time I run the program while testing, so it is commented out
        # .................................................................................
        # # Appends dataframe to existing sheet via mode='a'
        # with pd.ExcelWriter('Fall Schedule March 25th copy.xlsx', mode='a') as writer:
        #     df.to_excel(writer, sheet_name='DataFrame')
        # .................................................................................
