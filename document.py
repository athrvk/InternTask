import json
import pandas as pd
import time
import xlwings as xl

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
        # sh2.range('A1').value = ["City Name"]
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
        sh1.range('D' + city_index).value = city.temperature.is_temperature_c_or_f()
        sh1.range('E' + city_index).value = city.temperature.keep_updating

    print("Workbook initialized...")


def update_temperature(document, cities):
    sh1 = document.get_workbook().sheets("Sheet1")
    for city in cities:
        city_index = str(city.get_city_index() + 1)
        if city.temperature.keep_updating:
            current_time = time.ctime()
            temperature = city.get_temperature()

            sh1.range("B" + city_index).value = [current_time, '']
            time.sleep(0.25)

            sh1.range('C' + city_index).color = (0, 255, 0)
            sh1.range("C" + city_index).value = temperature
            print("Temperature updated for {}...".format(city.get_cityname()))
        else:
            sh1.range('C' + city_index).color = (128, 128, 128)
            print("Temperature not updated for {}...".format(city.get_cityname()))


def read_values_from_document(document, cities):
    sh1 = document.get_workbook().sheets("Sheet1")
    time.sleep(0.5)
    for city in cities:
        city_index = str(city.get_city_index() + 1)

        time.sleep(0.1)

        city.temperature.set_unit_system(
            set_metric_system(get_celld_value(city_index, sh1),
                              city_index)
        )

        city.temperature.set_update_flag(
            validate_update_flag(get_celle_value(city_index, sh1),
                                 city_index)
        )
    time.sleep(0.1)
    check_exit_flag(document, sh1)


def check_exit_flag(document, sh1):
    exit_flag = sh1.range('F2').value
    if exit_flag is not None and exit_flag == 2:
        print("Exiting program...")
        document.save_file()
        exit()


def get_celle_value(city_index, sh1):
    cell_e = sh1.range('E' + city_index).value
    return cell_e


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
    if not (cell_d == 'c' or cell_d == 'f'):
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
