import requests
import datetime
import openpyxl as px
import xlwings as xl
import json
import pandas as pd
import time


class intern_test:
    api_key = "7cac41f62c2d8231a68bbf5697ffea5a"

    check_update_flag = 0

    def __init__(self, city_name='Indore', unit_system='metric'):
        self.city_name = city_name
        self.unit_system = unit_system
        self.url = "http://api.openweathermap.org/data/2.5/weather?q={}&units={}&appid={}"

    def request_data(self):
        # print(self.url)
        requested_data = requests.get(self.url.format(self.city_name, self.unit_system, self.api_key))
        print("Data request complete...")
        return requested_data.json()

    def create_xlfile(self, filename='test.xlsx'):
        xl.Interactive = True
        xl.Visible = True
        workbook = xl.Book()
        workbook.activate(steal_focus=False)
        workbook.save(filename)
        print("Excel file : {} created...".format(filename))
        return workbook

    def load_xlfile(self, filename='test.xlsx'):
        print("Workbook loaded...")
        return xl.Book(filename)

    def initialize_workbook(self, workbook):
        sh1 = workbook.sheets("Sheet1")
        check = 'Note: Enter temperature option in cell "D2" and for stopping, in cell "E2" and press Enter Key'
        columns = ['              Date/Time', "City Name", "Temp", "Option C/F", 'Option to Stop Updating(0/1)', check]
        sh1.range('A1').value = columns
        sh1.autofit('c')
        sh2 = workbook.sheets.add(name="Sheet2", after=sh1)
        # sh2.range('A1').value = ["City Name"]
        self.write_citynames(workbook)
        print("Workbook initialized...")

    def append_data(self, data, workbook):
        dt = datetime.datetime.now()
        c_name = data['name']
        temp = data['main']['temp']
        # sheet1 = workbook.sheets['Sheet1']
        xl.Sheet('Sheet1').activate()
        # rng = xl.Range('A1').vertical.last_cell
        row = self.get_last_row(workbook, "Sheet1")
        # print(row)
        xl.Range("A" + str(row)).value = [dt, c_name, temp]
        print("Data appended successfully...")

    def get_last_row(self, workbook, sheet_name):
        row = workbook.sheets[sheet_name].range('A' + str(workbook.sheets['Sheet1'].cells.last_cell.row)).end(
            'up').row + 1
        return row

    def save_file(self, workbook, name='test.xlsx'):
        workbook.save()
        print("Workbook saved as {}...".format(name))

    def is_temperature_C_or_F(self, workbook):
        sh1 = workbook.sheets['Sheet1']
        option = ''

        if sh1.range('D2').value is not None:
            option = str(sh1.range("D2").value.lower())
            # print(option)
        if option == 'c':
            self.unit_system = 'metric'
            # print(self.unit_system)
        if option == 'f':
            self.unit_system = 'imperial'
            # print(self.unit_system)

    def to_stop_updating(self, workbook):
        sh1 = workbook.sheets['Sheet1']
        option = ''

        if sh1.range('E2').value is not None:
            option = int(sh1.range('E2').value)
            # print(option)
        if option == 1:
            self.check_update_flag = 1
            # print(self.flag)
        if option == 0:
            self.check_update_flag = 0
            # print(self.flag)

    def write_citynames(self, workbook):
        sh2 = workbook.sheets['Sheet2']
        with open('city.list.json', encoding='utf8') as fp:
            c = json.load(fp)

            df = pd.DataFrame(c)
            df['City Name'] = df['name']
            sh2.range("A1").options(index=False, headers=False).value = df['City Name']
            sh2.autofit('c')
        time.sleep(1.5)
