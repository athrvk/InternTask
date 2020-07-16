from temperature import Temperature


class City:

    def __init__(self, name, index):
        self.__name = name
        self.__index = index
        self.temperature = Temperature(self.__name)

    def get_cityname(self):
        return self.__name

    def get_city_index(self):
        return self.__index

    def get_temperature(self):
        self.temperature.fetch_temperature()
        return self.temperature.get_temperature()
