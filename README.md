# InternTask

This project uses a live weather api and updates the temperature of user given city at a regular interval. 
New cities can be added by entering it in the next empty cell in column A 

> The program works under Windows only(tested in Windows 10, conda environment, PyCharm IDE) and requires Microsoft Excel to be installed on the system.

all the necessary files are uploaded

The program uses 'xlwings' to handle the excel instance

you can use the following command to set up the environment

`conda env create --file environment.yml`

then run the use_api.py file

`python start.py`


the list of cities is gathered from `https://bulk.openweathermap.org/sample/`
