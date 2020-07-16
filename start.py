from document import *
from city import *

document = Document(filename="weather.xlsx")
indore = City("Indore", 1)
bhopal = City("Bhopal", 2)
noida = City('Noida', 3)
mumbai = City('Mumbai', 4)
delhi = City('Delhi', 5)

cities = [indore, bhopal, noida, mumbai, delhi]


def start_the_program():
    initialize_document(document, cities)
    write_citynames(document)  # in Sheet2

    while True:
        read_values_from_document(document, cities)
        update_temperature(document, cities)


if __name__ == '__main__':
    start_the_program()
