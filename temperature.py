import requests


class Temperature:

    def __init__(self, cityname, unit_system="metric"):
        self.__temperature = float
        self.__city = cityname
        self.keep_updating = 1
        self.__unit_system = unit_system

    def fetch_temperature(self):
        global api_return_value
        api_key = "7cac41f62c2d8231a68bbf5697ffea5a"
        url = "http://api.openweathermap.org/data/2.5/weather?q={}&units={}&appid={}"
        try:
            api_return_value = requests.get(url.format(self.__city, self.__unit_system, api_key))
            if api_return_value.json()['cod'] == '404':
                print('Wrong City Name: Enter valid city name refer Sheet 2...')
                exit()
        except Exception as e:
            print(e, "\nCannot Request Data...\n")
            exit()

        api_return_value_json = api_return_value.json()
        self.__temperature = api_return_value_json['main']['temp']

    def get_temperature(self):
        return self.__temperature

    def set_unit_system(self, value):
        self.__unit_system = value

    def is_temperature_c_or_f(self):
        if self.__unit_system == 'metric':
            return 'C'
        if self.__unit_system == 'imperial':
            return 'F'

    def set_update_flag(self, value):
        self.keep_updating = value
