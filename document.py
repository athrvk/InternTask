import json
import pandas as pd
import time
import xlwings as xl
from city import *
import threading
import _thread

# Constants
F = 'f'
C = 'c'
IMPERIAL = 'imperial'
METRIC = 'metric'
update_true = 1
update_false = 0


# Load already created Excel File
def load_xlfile(filename):
    try:
        print("Workbook loaded...")
        return xl.Book(filename)
    except Exception as e:
        print("Error: File does not exist...")
        exit()


# Write City Names to Sheet2
def write_citynames(document):
    sh2 = document.get_workbook().sheets['Sheet2']
    with open('city_list.json', encoding='utf8') as fp:
        c = json.load(fp)
        df = pd.DataFrame(c)
        df['City Name'] = df['name']
        sh2.range("A1").options(index=False, headers=False).value = df['City Name']
        sh2.autofit('c')
    time.sleep(0.1)


# Initialize the document created i.e, adding headers, city name in Sheet2
def initialize_document(document, cities):
    try:
        sh2 = document.get_workbook().sheets.add(name="Sheet2", after='Sheet1')
        write_citynames(document)
        print("City Names added in Sheet2...")
        time.sleep(0.1)
    except Exception as e:
        print(e, "\nInstance of the file already open in Excel. Close Excel and re-run the program...")
        exit()

    xl.Sheet('Sheet1').activate()
    sh1 = document.get_workbook().sheets("Sheet1")
    check = 'Note: To stop the program Enter 2 in F2'
    columns = ["City Name", 'Last Updated                         ',
               "Temp", "Option C/F", 'Update Temperature(0/1)', check]
    sh1.range('A1').value = columns
    sh1.autofit('c')

    for city in cities:
        city_index = str(city.get_city_index() + 1)
        sh1.range('A' + city_index).value = city.get_cityname()
        sh1.range('D' + city_index).value = [city.temperature.is_temperature_c_or_f(), city.temperature.keep_updating]

    print("Workbook initialized...")

threadLock = threading.Lock()
def update_temperature_threaded(cities):
    sh1 = xl.Book('weather.xlsx').sheets['Sheet1']
    while True:
        print("Updating")
        for city in cities:
            city_index = str(city.get_city_index() + 1)
            threadLock.acquire()
            if city.temperature.keep_updating:
                sh1.range("B" + city_index).value = [time.ctime(), '']
                time.sleep(0.2)

                sh1.range('C' + city_index).color = (0, 255, 0)
                sh1.range("C" + city_index).value = city.get_temperature()
                print("Temperature updated for {}...".format(city.get_cityname()))
            else:
                sh1.range('C' + city_index).color = (128, 128, 128)
                print("Temperature not updated for {}...".format(city.get_cityname()))
            threadLock.release()


def update_temperature(document, cities):
    sh1 = document.get_workbook().sheets("Sheet1")
    for city in cities:
        city_index = str(city.get_city_index() + 1)
        if city.temperature.keep_updating:
            sh1.range("B" + city_index).value = [time.ctime(), '']
            time.sleep(0.2)

            sh1.range('C' + city_index).color = (0, 255, 0)
            sh1.range("C" + city_index).value = city.get_temperature()
            print("Temperature updated for {}...".format(city.get_cityname()))
        else:
            sh1.range('C' + city_index).color = (128, 128, 128)
            print("Temperature not updated for {}...".format(city.get_cityname()))


def read_values_threaded(cities):
    sh1 = xl.Book('weather.xlsx').sheets['Sheet1']
    time.sleep(0.1)
    while True:
        threadLock.acquire()
        print("Reading")
        input_values = sh1.range('D2:E{}'.format(len(cities) + 1)).value

        for unit_system, update_flag in input_values:
            if not ((unit_system.lower() != 'c') ^ (unit_system.lower() != 'f')):
                print('Invalid Entry: Enter either C or F')
                raise ValueError()
            if not ((update_flag != 1) ^ (update_flag != 0)):
                print('Invalid Entry: Enter either 0 or 1')
                raise ValueError()

        for index in range(len(cities)):
            cities[index].temperature.set_update_flag(input_values[index][1])
            cities[index].temperature.set_unit_system(conversion(input_values[index][0].lower()))

        threadLock.release()
        check_exit_flag(xl.Book('weather.xlsx'), sh1)
        check_new_city(cities, sh1)
        # time.sleep(2)


def read_values_from_document_v2(document, cities):
    sh1 = document.get_workbook().sheets("Sheet1")
    time.sleep(0.5)

    input_values = sh1.range('D2:E{}'.format(len(cities) + 1)).value

    for unit_system, update_flag in input_values:
        if not ((unit_system.lower() != 'c') ^ (unit_system.lower() != 'f')):
            print('Invalid Entry: Enter either C or F')
            raise ValueError()
        if not ((update_flag != 1) ^ (update_flag != 0)):
            print('Invalid Entry: Enter either 0 or 1')
            raise ValueError()

    for index in range(len(cities)):
        cities[index].temperature.set_update_flag(input_values[index][1])
        cities[index].temperature.set_unit_system(conversion(input_values[index][0].lower()))

    check_exit_flag(document, sh1)

    check_new_city(cities, sh1)


def read_values_from_document(document, cities):
    sh1 = document.get_workbook().sheets("Sheet1")
    time.sleep(0.5)
    for city in cities:
        city_index = str(city.get_city_index() + 1)
        city.temperature.set_unit_system(
            set_metric_system(get_celld_value(city_index, sh1),
                              city_index)
        )
        # time.sleep(0.1)
        city.temperature.set_update_flag(
            validate_update_flag(get_celle_value(city_index, sh1),
                                 city_index)
        )
    # time.sleep(0.1)
    check_exit_flag(document, sh1)

    check_new_city(cities, sh1)


def check_new_city(cities, sh1):
    next_empty_cell = len(cities) + 2
    cell_value = sh1.range('A' + str(next_empty_cell)).value
    if cell_value is not None:
        print(cell_value)
        new_city = City(cell_value, next_empty_cell - 1)
        cities.append(new_city)
        sh1.range('D' + str(next_empty_cell)).value = ['C', 1]


def check_exit_flag(document, sh1):
    exit_flag = sh1.range('F2').value
    if exit_flag is not None and exit_flag == 2:
        print("Exiting program...")
        document.save()
        # document.save_file()
        # exit()
        _thread.interrupt_main()
    return True


def get_celle_value(city_index, sh1):
    return sh1.range('E' + city_index).value


def get_celld_value(city_index, sh1):
    cell_d = str(sh1.range('D' + city_index).value)
    cell_d = cell_d.lower()
    return cell_d


def validate_update_flag(cell_e, city_index):
    if not is_valid_flag(cell_e):
        print("Invalid Entry(cell E{}): Enter either 0 or 1\n"
              "Restart the program".format(city_index))
        exit()

    return cell_e


def is_valid_flag(cell_e):
    return cell_e == update_false or cell_e == update_true


def set_metric_system(cell_d, city_index):
    if not (cell_d == C or cell_d == F):
        print("Invalid Entry(cell D{}): Enter either 'C' or 'F'\n"
              "Restart the program".format(city_index))
        exit()
    return conversion(cell_d)


def conversion(value):
    if value == 'c':
        return METRIC
    elif value == 'f':
        return IMPERIAL


class Document:

    def __init__(self, filename='default.xlsx'):
        self.__filename = filename
        self.__workbook = self.create_xlfile()

    def get_filename(self):
        return self.__filename

    def create_xlfile(self):
        global temp_workbook
        xl.Interactive = True
        xl.Visible = True
        try:
            temp_workbook = xl.Book()
            temp_workbook.activate(steal_focus=True)
            temp_workbook.save(self.__filename)
        except Exception:
            print("\nExcel not installed or file already opened or file cannot be created...")
            exit()

        print("Excel file : {} created...".format(self.__filename))
        return temp_workbook

    def get_workbook(self):
        return self.__workbook

    def save_file(self):
        self.__workbook.save()
        print("Workbook saved...")
