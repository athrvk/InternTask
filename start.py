from document import *
from city import *

document = Document(filename="weather.xlsx")

cities = [City("Indore", 1), City("Bhopal", 2), City('Noida', 3), City('Mumbai', 4), City('Delhi', 5)]


def start_the_program():
    initialize_document(document, cities)
    write_citynames(document)  # in Sheet2

    while True:
        read_values_from_document(document, cities)
        update_temperature(document, cities)


if __name__ == '__main__':
    start_the_program()
