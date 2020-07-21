# InternTask

This project uses a live weather api and updates the temperature of user given city at a regular interval. 

## Features
-> Implemented on the concept of Object Oriented Programing (Classes, Objects)

-> New cities can be added dynamically by entering it in the next empty cell in column A 

-> Error handling for invalid inputs(with proper error message), multiple instance of excel

## Limitations
-> A city cannot be removed once entered in the cell

-> xlwings library does not have event listners, so handling input events is hardcoded

-> Excel freezes* if the program read/writes while the user also enters the value in a cell

*(for random period, on average 3 seconds; but restores itself back and continues to run without any error)


> The program works under Windows only(tested in Windows 10, conda environment, PyCharm IDE) and requires Microsoft Excel to be installed on the system.



The program uses 'xlwings' to handle the excel instance

you can use the following command to set up the environment

`conda env create --file environment.yml`

then run the start.py file

`python start.py`


the list of cities is gathered from `https://bulk.openweathermap.org/sample/`
