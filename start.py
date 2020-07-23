from document import *
from city import *
import threading

document = Document(filename="weather.xlsx")

cities = [City("Indore", 1), City("Bhopal", 2),
          City('Noida', 3), City('Mumbai', 4),
          City('Delhi', 5)]

global t


def start_the_program():
    initialize_document(document, cities)

    # while True:
        # read_values_from_document(document, cities)
        # read_values_from_document_v2(document, cities)
        # update_temperature(document, cities)

    r = threading.Thread(target=read_values_threaded, args=[cities])
    u = threading.Thread(target=update_temperature_threaded, args=[cities], daemon=True)
    r.start()
    u.start()
    r.join()
    u.join()


if __name__ == '__main__':
    start_the_program()
