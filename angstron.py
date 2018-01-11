import xlrd
import math


def angstronIndex (airTemperature, relactiveHumidity):
    ###This function calculates the Angstron Index for a given air temperature and relative humidity, measured at 12h.
    angstron = 0.05 * relactiveHumidity - 0.1 * (airTemperature - 27)
    print('Angstron Index is:', round(angstron, 2))


def telicynIndex(airTemperature, dewPoint, N):
    ###This function calculates the Telicyn Index for a set of air temperature and dew point, measured at 13h.
    i = 0
    telicyn = 0.0
    while i < N:
        telicyn = math.log(airTemperature - dewPoint, 10) + telicyn
        i = i + 1
    print('Telicyn index is:', round(telicyn, 2))

def nesterovIndex(airTemperature, airDeficit, N):
    ###This function calculates the Nesterov Index for a set of air temperature and air deficit, measured at 13h.
    i = 0
    nesterov = 0.0
    while i < N:
        #falta calcular o airDeficit
        nesterov = airDeficit - airTemperature + nesterov

    print('Nesterov Index is:', round(nesterov,2))

def monteAlegre(relactiveHumidity, N):
        ###This function calculates the Monte Alegre's Formula for a set of relative humidity, measured at 13h.
    i = 0
    monte = 0.0
    while i < N:
        monte = (100 / relactiveHumidity) + monte
        i = i + 1
    print('Monte Alegre Formula is:', round(monte,2))


print("""\t\tHi! This is the Index Incendies Calculator - IIC.
Please follow the below instructions in order to calculate one of the available indexes.\n\n""")
xls_path = input("""Tell us where the .XLS (Excel File) is located in your computer.
Remeber to use /home/user/Documents/<archieve_name>.xls for Linux or C:/Documents/<archieve_name>.xls for Windows\n""")

workbook = xlrd.open_workbook(xls_path)
worksheet = workbook.sheet_by_index(0)
option = 0
menu = True

while menu == True:
    option = input("""\n\n This is the IIC menu. Type the number to use the following formula.
    1 - Angstron Index
    2 - Telicyn Index
    3 - Nesterov Index
    4 - Monte Alegre Formula
    5 - Quit\n""")

    if option == '1':
        airTemp_line = int(input("\nWhat's the line on the Excel file that contains the Air Temperature information?\n Example: If the air temperature is at the line 13 on the excel document, you must use the value of 12. (1->13 = 12)\n"))
        airTemp_col = int(input("\nWhat's the column on the Excel file that contains the Air Temperature information?\n Example: If the air temperature is at the column N on the excel document, you must use the value of 13. (1->14(N) = 13)\n"))
        relHum_line = int(input("\nWhat's the column on the Excel file that contains the Relative Humidity information?\n Example: If the relative humidity is at the line 13 on the excel document, you must use the value of 12. (1->13 = 12)\n"))
        relHum_col = int(input("\nWhat's the column on the Excel file that contains the Relative Humidity information?\n Example: If the relative humidity is at the column AL on the excel document, you must use the value of 37. (1->38(AL) = 37)\n"))

        airTemp_value = worksheet.cell(airTemp_line, airTemp_col).value
        relHum_value = worksheet.cell(relHum_line, relHum_col).value
        angstronIndex(airTemp_value, relHum_value)

        
    elif option == '2':
        # calculates telicyn index
        print("something")


    elif option == '3':
        # calculates nesterov index
        print("something")


    elif option == '4':
        # calculates monte carlo formula
        print("something")


    elif option == '5':
        menu = False

    else:
        print("The input value is invalid. The program is finishing.")
        menu = False
#airTemp12 = worksheet.cell(12, 13).value
#umiRel = worksheet.cell(12, 37).value

#airTemp13 = worksheet.cell(12, 14).value
#dewPt = worksheet.cell(12, 62).value

#umiRel13 = worksheet.cell(12, 38).value

#angstronIndex(airTemp12, umiRel)
#telicynIndex(airTemp13, dewPt, 1)
#nesterofIndex(airTemp13, )
#monteAlegre(umiRel13, 1)
