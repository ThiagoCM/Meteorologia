import xlrd
import math
def angstronIndex (airTemperature, relactiveHumidity):
    ###This function calculates the Angstron Index for a given air temperature and relactive humidity, , measured at 12h.
    result = 0.05 * relactiveHumidity - 0.1 * (airTemperature - 27)
    print('Angstron Index is:', round(result, 2))


def telicynIndex(airTemperature, dewPoint, N):
    ###This function calculates the Telicyn Index for a set of air temperature and dew point, measured at 13h.
    i = 0
    result = 0.0
    while i < N:
        result = math.log(airTemperature - dewPoint, 10) + result
        i = i + 1
    print('Telicyn index is:', round(result, 2))

def nesterovIndex(airTemperature, airDeficit, N):
    ###This function calculates the Nesterov Index for a set of air temperature and air deficit, measured at 13h.
    i = 0
    result = 0.0
    while i < N:
        #falta calcular o airDeficit
        result = airDeficit - airTemperature + result

    print('Nesterov Index is:', result)

def monteAlegre(relactiveHumidity, N):
        ###This function calculates the Monte Alegre's Formula for a set of relactive humidity, measured at 13h.
    i = 0
    result = 0.0
    while i < N:
        result = (100 / relactiveHumidity) + result
        i = i + 1
    print('Monte Alegre Formula is:', round(result,2))

print("""\t\tHi! This is the Index Incendies Calculator - IIC.
Please follow the below instructions in order to calculate one of the available indexes.\n\n""")
xls_path = raw_input("""Tell us where the .XLS (Excel File) is located in your computer.
Remeber to use /home/user/Documents/<archieve_name>.xls for Linux or C:/Documents/<archieve_name>.xls for Windows\n""")

workbook = xlrd.open_workbook(xls_path)
worksheet = workbook.sheet_by_index(0)
option = 0

while True:
    option = int(raw_input("""\n\n This is the IIC menu. Type the number to calculate it's index.
    1 - Angstron Index
    2 - Telicyn Index
    3 - Nesterov Index
    4 - Monte Alegre Formula\n"""))

    if option == 1:
        # calculates angstron index
        
    elif option == 2:
        # calculates telicyn index

    elif option == 3:
        # calculates nesterov index

    elif option == 4:
        # calculates monte carlo formula

airTemp12 = worksheet.cell(12, 13).value
umiRel = worksheet.cell(12, 37).value

airTemp13 = worksheet.cell(12, 14).value
dewPt = worksheet.cell(12, 62).value

umiRel13 = worksheet.cell(12, 38).value

angstronIndex(airTemp12, umiRel)
telicynIndex(airTemp13, dewPt, 1)
#nesterofIndex(airTemp13, )
monteAlegre(umiRel13, 1)
