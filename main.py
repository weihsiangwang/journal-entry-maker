import os
import pandas as pd
from datetime import datetime
import string
from xlsxwriter.utility import xl_range
import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename, askdirectory
from tkinter import messagebox
from tkcalendar import DateEntry


class App:
    def __init__(self, parent):
        # variables
        self.insert_type = tk.StringVar()
        self.df = None
        self.voucher_list = None
        self.date_list = None
        self.input_list = tk.StringVar()

        self.parent = parent
        self.parent.geometry('300x320')
        self.parent.title("Journal Entry GUI")

        #### FILE FRAME ####
        file_frame = tk.LabelFrame(self.parent)
        file_frame.grid(row=0, column=0, columnspan=3,
                        sticky='nsew', padx=10, pady=5)

        self.file_text = tk.Label(file_frame, text="SOURCE FILE", font=(
            "Arial Bold", 10), width=20, anchor='w')
        self.file_text.grid(row=0, column=0, sticky="w", padx=5)
        self.filename = tk.StringVar()
        self.filename.set(r'Journal-2022.xlsx')
        self.file_path_label = tk.Label(
            file_frame, textvariable=self.filename, width=20, borderwidth=1, relief="sunken", anchor="w")
        self.file_path_label.grid(row=1, column=0, sticky="w", padx=5)
        self.file_button = tk.Button(
            file_frame, text='LOAD DATA', command=self.load, width=10)
        self.file_button.grid(row=0, column=1, sticky="w", padx=5)

        self.load_file_msg = tk.StringVar()
        self.load_file_msg.set(' ')
        self.file_msg = tk.Label(
            file_frame, textvariable=self.load_file_msg, anchor="w")
        self.file_msg.grid(row=2, column=0, sticky="w", padx=5)

        #### INSERT TYPE ####
        self.type_frame = tk.Frame(self.parent)
        self.type_frame.grid(row=1, column=0, columnspan=3,
                             sticky='nsew', padx=10, pady=(0, 5))
        tk.Label(self.type_frame, text="INSERT TYPE", font=(
            "Arial Bold", 10), anchor='w', justify='left').pack(side='left', padx=5)
        self.insert_type.set('NUMBER')
        tk.Radiobutton(self.type_frame, variable=self.insert_type,
                       text='Number', value='NUMBER', command=self.enable_entry).pack(side='left')
        tk.Radiobutton(self.type_frame, variable=self.insert_type,
                       text='Date', value='DATE', command=self.enable_dateentry).pack(side='left')

        #### VOUCHER TITLE ####
        self.input_text = tk.Label(
            self.parent, text="VOUCHER NUMBER", font=("Arial Bold", 10), anchor='w')
        self.input_text.grid(
            row=2, column=0, columnspan=3, sticky="w", padx=15)
        self.insert_type.trace('w', self.change_input_type_label)

        #### NUMBER FRAME ####
        number_frame = tk.Frame(self.parent)
        number_frame.grid(row=3, column=0, sticky='nsew', padx=(10, 0), pady=5)
        vcmd = number_frame.register(self.vcmdDigital)
        self.input = tk.Entry(number_frame, width=13,
                              validate='all', validatecommand=(vcmd, '%P'))
        self.input.grid(row=0, column=0, padx=5, pady=5)
        select_date = tk.StringVar()
        self.cal = DateEntry(number_frame, width=10, selectmode='day',
                             textvariable=select_date, date_pattern='yyyy/mm/dd')
        self.input_data_msg = tk.StringVar()
        self.input_data_msg.set(' ')
        self.input_msg = tk.Label(
            number_frame, textvariable=self.input_data_msg, anchor="w")
        self.input_msg.grid(row=1, column=0, padx=5)
        self.input_button = tk.Button(
            number_frame, text='INSERT', command=self.insert, width=8)
        self.input_button.grid(row=2, column=0, padx=5, pady=5)
        self.input_button = tk.Button(
            number_frame, text='DELETE', command=self.delete, width=8)
        self.input_button.grid(row=3, column=0, padx=5, pady=5)

        #### LISTBOX FRAME ####
        listbox_frame = tk.Frame(self.parent)
        listbox_frame.grid(row=3, column=1, sticky='nsew', pady=5)
        self.scrollbar = tk.Scrollbar(listbox_frame)
        self.scrollbar.pack(side="right", fill="y")
        self.listbox = tk.Listbox(
            listbox_frame, listvariable=self.input_list, width=10, height=7, yscrollcommand=self.scrollbar.set)
        self.listbox.pack(side="left", fill="both")
        self.scrollbar.config(command=self.listbox.yview)

        #### REPORT FRAME ####
        self.report_button = tk.Button(
            self.parent, text="Report", width=10, height=2, bg="orange", fg="red", command=self.report)
        self.report_button.grid(row=3, column=2, padx=10, pady=15)
        self.report_button_msg = tk.StringVar()
        self.report_button_msg.set(' ')

        #### MESSAGE FRAME ####
        self.report_msg = tk.Label(
            self.parent, textvariable=self.report_button_msg, wraplength=250, justify="center", anchor="w")
        self.report_msg.grid(row=4, column=0, columnspan=3)

    def change_input_type_label(self, *args):
        title = 'VOUCHER '
        title = title + self.insert_type.get()
        self.input_text['text'] = title

    def enable_entry(self):
        # enable input number entry
        self.input.grid(row=0, column=0, padx=5, pady=5)
        # disable calendar
        self.cal.grid_forget()
        # clear list box
        self.listbox.delete(0, 'end')
        self.input_data_msg.set(' ')
        self.report_button_msg.set(' ')

    def enable_dateentry(self):
        # disable input number entry
        self.input.grid_forget()
        # enable calendar
        self.cal.grid(row=0, column=0, padx=5, pady=5)
        # clear list box
        self.listbox.delete(0, 'end')
        self.input_data_msg.set(' ')
        self.report_button_msg.set(' ')

    def split_file_name(self, path):
        slash_index = max(path.rfind('\\'), path.rfind('/'))
        if (slash_index > 0):
            folder = path[:slash_index+1]
            filename = path[slash_index+1:]
        else:
            folder = ''
            filename = path
        return folder, filename

    def load(self):
        # load file
        current_path = os.getcwd()
        file_path = askopenfilename(initialdir=current_path, title="Choose your journal entry data...", filetypes=[
                                    ('Excel Files', ('*.xls', '*.xlsx')), ('CSV Files', '*.csv',)])
        if file_path:
            if file_path.endswith('.csv'):
                data = pd.read_csv(file_path, dtype=str)
            else:
                data = pd.read_excel(file_path, dtype=str)

            # convert columns data type
            column_name = data.columns
            if set(['日期', '傳票號碼', '會計項目', '項目名稱', '摘要', '借方金額', '貸方金額']).issubset(column_name):
                data['日期'] = pd.to_datetime(data['日期'], format='%Y-%m-%d')
                data['借方金額'] = data['借方金額'].apply(float)
                data['貸方金額'] = data['貸方金額'].apply(float)
                self.voucher_list = data['傳票號碼'].unique().tolist()
                self.date_list = data['日期'].dt.strftime(
                    '%Y/%m/%d').unique().tolist()
                self.df = data
            else:
                messagebox.showerror(
                    "Error", "Please check the column names!\n needed: ['日期', '傳票號碼', '會計項目', '項目名稱', '摘要', '借方金額', '貸方金額']")

            file_folder, file_name = self.split_file_name(file_path)
            self.filename.set(file_name)
            self.load_file_msg.set(' ... File loaded!')
        else:
            self.filename.set(' ')
            self.load_file_msg.set('Please check your file selected...')

    def save(self):
        current_path = os.getcwd()
        dic_path = askdirectory(initialdir=current_path,
                                title="Choose the output directory...")
        return dic_path

    # Restricting the value entry
    def vcmdDigital(self, P):
        if str.isdigit(P) or str(P) == "":
            return True
        else:
            return False

    def insert(self):
        input_type = self.insert_type.get()
        if input_type == 'NUMBER':
            input_value = self.input.get()
        elif input_type == 'DATE':
            input_value = self.cal.get()

        # check if input exists
        if self.voucher_list is None or self.date_list is None:
            messagebox.showerror("Error", "Load file first!")
            self.load()
        elif len(input_value) == 0:
            messagebox.showwarning("Warning", "Type a value first!")
        elif input_value in self.listbox.get(0, 'end'):
            self.input_data_msg.set('exist!')
        elif input_type == 'NUMBER' and input_value in self.voucher_list:
            self.listbox.insert('end', input_value)
            self.input_data_msg.set('inserted')
        elif input_type == 'DATE' and input_value in self.date_list:
            self.listbox.insert('end', input_value)
            self.input_data_msg.set('inserted')
        else:
            messagebox.showerror("Error", "Value does not exist!")

    def delete(self):
        if self.listbox.curselection():
            select_index = self.listbox.curselection()
            print(select_index)
            self.listbox.delete(select_index)

    # cell formatting
    def workbookFormat(self, wbk):
        wbk_fmt = dict()
        wbk_fmt['header'] = wbk.add_format({
            'font_name': '標楷體',
            'font_size': '18',
            'align': 'center'})
        wbk_fmt['colname'] = wbk.add_format({
            'font_name': '標楷體',
            'font_size': '18',
            'bold': True,
            'align': 'center'})
        wbk_fmt['num'] = wbk.add_format({
            'font_name': '標楷體',
            'font_size': '18',
            'num_format': '#,##0'})
        wbk_fmt['date'] = wbk.add_format({
            'font_name': '標楷體',
            'font_size': '18',
            'num_format': 'yyyy/mm/dd'})
        wbk_fmt['default'] = wbk.add_format({
            'font_name': '標楷體',
            'font_size': '18'})
        wbk_fmt['border'] = wbk.add_format({
            'bottom': 1,
            'top': 1,
            'left': 1,
            'right': 1})
        wbk_fmt['underline'] = wbk.add_format({
            'bottom': 5})

        return wbk_fmt

    def report(self):
        # check if data loaded
        temp_list = self.listbox.get(0, 'end')
        if self.df is None:
            messagebox.showerror("Error", "Load file first!")
            self.load()
            return
        elif len(temp_list) == 0:
            messagebox.showwarning("Warning", "Insert a number at least!")
            return
        # using list by selected type
        input_type = self.insert_type.get()
        if input_type == 'NUMBER':
            number_list = temp_list
        elif input_type == 'DATE':
            date_list = temp_list
            temp_df = self.df[self.df['日期'].isin(date_list)]
            number_list = temp_df['傳票號碼'].unique().tolist()
        print(number_list)

        # GO
        timestamp = datetime.today().strftime('%Y%m%d%H%M%S')
        report_file_folder = self.save()
        report_file_name = report_file_folder + '\\Account Entry_by' + \
            input_type + '_' + timestamp + '.xlsx'

        num_format_columns = ['借方金額', '貸方金額']
        account_columns_name = ['會計項目', '項目名稱', '摘要', '借方金額', '貸方金額']
        account_column_width = [20, 20, 60, 20, 20]
        # get 26 upper letters to define columns
        alph_up = string.ascii_letters[26:]

        with pd.ExcelWriter(report_file_name, engine='xlsxwriter') as writer:
            workbook = writer.book
            workbook_format = self.workbookFormat(workbook)

            for voucher in number_list:
                # entry
                df_writer = self.df.loc[(self.df['傳票號碼'] == voucher) & (
                    self.df['會計項目'] != '合計')].reset_index()
                df_account = df_writer[account_columns_name]
                df_date = df_writer.日期[0]
                # total
                df_total = self.df.loc[(self.df['傳票號碼'] == voucher) & (
                    self.df['會計項目'] == '合計')].reset_index()
                df_total = df_total[account_columns_name]
                # turn off the default header and skip one row
                df_account.to_excel(writer, sheet_name=voucher,
                                    startrow=6, index=False, header=False)

                worksheet = writer.sheets[voucher]
                for col_num, col_name in enumerate(df_account.columns.values):
                    alph = '{}:{}'.format(alph_up[col_num], alph_up[col_num])
                    if col_name in num_format_columns:
                        worksheet.set_column(
                            alph, cell_format=workbook_format['num'], width=account_column_width[col_num])
                    else:
                        worksheet.set_column(
                            alph, cell_format=workbook_format['default'], width=account_column_width[col_num])
                    # Write the column headers with the defined format
                    worksheet.write_string(
                        row=5, col=col_num, string=col_name, cell_format=workbook_format['colname'])

                # header
                worksheet.merge_range('A1:E1', 'Company',
                                      workbook_format['header'])
                worksheet.write_string(
                    'C2', '轉  帳  傳  票', workbook_format['header'])
                # Voucher Date
                worksheet.write_string(
                    'A4', '傳票日期', workbook_format['default'])
                worksheet.write_datetime(
                    'B4', df_date, workbook_format['date'])
                # department
                worksheet.write_string('A5', '部門：', workbook_format['default'])
                # Voucher Number
                worksheet.write_string(
                    'D5', '傳票編號：', workbook_format['default'])
                worksheet.write_string(
                    'E5', voucher, workbook_format['default'])
                # border:
                # row 5~ , col 0~
                report_row_num = (len(df_account)//10+1)*10
                report_col_num = len(df_account.columns)
                worksheet.conditional_format(xl_range(5, 0, 5+report_row_num, report_col_num-1), {
                                             'type': 'no_errors', 'format': workbook_format['border']})
                worksheet.conditional_format(
                    'C2', {'type': 'no_errors', 'format': workbook_format['underline']})
                # total
                df_total.to_excel(
                    writer, sheet_name=voucher, startrow=5+report_row_num, index=False, header=False)
                alph = 'A{}:C{}'.format(5+report_row_num+1, 5+report_row_num+1)
                worksheet.merge_range(
                    alph, '合         計', workbook_format['header'])
                # footer
                alph = 'A{}:E{}'.format(5+report_row_num+2, 5+report_row_num+2)
                worksheet.merge_range(
                    alph, '核准：           覆核：           會計：           出納：', workbook_format['default'])
                # 1 page wide and as long as necessary.
                worksheet.fit_to_pages(1, 0)
                # Set the number of rows to repeat at the top of each printed page.
                worksheet.repeat_rows(first_row=0, last_row=5)
                # page break row
                page_break_list = [
                    6+(i+1)*10 for i in range(len(df_account)//10)]
                worksheet.set_h_pagebreaks(page_break_list)

        self.report_button_msg.set('Report Saved! ... \n' + report_file_folder)


if __name__ == '__main__':
    root = tk.Tk()
    top = App(root)
    root.iconbitmap(".\img\libraries.ico")
    root.mainloop()

# pyinstaller -F -w --hidden-import "babel.numbers" -i "libraries.ico" "Journal Entry GUI.py"
