import xlrd
import math
import matplotlib.pyplot as plt

def ploting1(dailyValues, indexes):
    if(indexes == [1,2.5,5,7.5,10]):
        plt.plot(1,dailyValues, 'ro')
        plt.axis([0,10,0,7.5], 2)
        plt.ylabel('Index Values for Angstron Index')
        plt.xlabel('Angstron Daily Values')
        plt.show()
    if(indexes == [2, 3.5, 5]):
        plt.plot(1,dailyValues, 'ro')
        plt.axis([0,10,0,10], 2)
        plt.ylabel('Index Values for Telicyn Index')
        plt.xlabel('Telicyn Daily Values')
        plt.show()
    if(indexes == [300, 500, 1000, 4000]):
        plt.plot(1,dailyValues, 'ro')
        plt.axis([0,10,0,4000], 2)
        plt.ylabel('Index Values for Nesterov Index')
        plt.xlabel('Nesterov Daily Values')
        plt.show()
    if(indexes == [1, 3, 8, 20]):
        plt.plot(1,dailyValues, 'ro')
        plt.axis([0,10,0,40], 2)
        plt.ylabel('Index Values for Monte Alegre Index')
        plt.xlabel('Monte Alegre Daily Values')
        plt.show()


def angstronIndex (airTemperature, relactiveHumidity):
    ###This function calculates the Angstron Index for a given air temperature and relative humidity, measured at 12h.
    angstron = 0.05 * relactiveHumidity - 0.1 * (airTemperature - 27)
    print('Angstron Index is:', round(angstron, 2))
    ploting1(angstron, [1,2.5,5,7.5,10])


def telicynIndex(airTemperature, dewPoint, N):
    ###This function calculates the Telicyn Index for a set of air temperature and dew point, measured at 13h.
    i = 0
    telicyn = 0.0
    while i <= N:
        if(airTemperature[i] == 'NULL' or dewPoint[i] == 'NULL'):
            i = i + 1
            continue
        telicyn = math.log(airTemperature[i] - dewPoint[i], 10) + telicyn
        i = i + 1
    print('Telicyn index is:', round(telicyn, 2))
    ploting1(telicyn, [2, 3.5, 5])


def nesterovIndex(airTemperature, relactiveHumidity, N):
    ###This function calculates the Nesterov Index for a set of air temperature and air deficit, measured at 13h.
    i = 0
    nesterov = 0.0
    while i < N:
        pressure = 0,6108 * math.pow(10, ((7,5 * airTemperature)/(237,5+airTemperature)))
        partialPressure = relactiveHumidity * pressure / 100
        pressureDeficit = pressure - partialPressure
        nesterov = airDeficit - airTemperature + nesterov

    print('Nesterov Index is:', round(nesterov,2))
    ploting1(nesterov, [300, 500, 1000, 4000])

def monteAlegre(relactiveHumidity, N):
        ###This function calculates the Monte Alegre's Formula for a set of relative humidity, measured at 13h.
    i = 0
    monte = 0.0
    while i <= N:
        if(relactiveHumidity[i] == 'NULL'):
            i = i + 1
            continue
        monte = (100 / relactiveHumidity[i]) + monte
        i = i + 1
    print('Monte Alegre Formula is:', round(monte,2))
    ploting1(monte, [1, 3, 8, 20])



print("""\t\tHi! This is the Index Incendies Calculator - IIC.
Please follow the below instructions in order to calculate one of the available indexes.\n\n""")
xls_path = input("""Tell us where the .XLS (Excel File) is located in your computer.
Remeber to use /home/user/Documents/<archieve_name>.xls for Linux or C:/Documents/<archieve_name>.xls for Windows\n""")

workbook = xlrd.open_workbook(xls_path)
worksheet = workbook.sheet_by_index(0)
option = 0
menu = True

while menu == True:
    index = input("""\n\n This is the IIC menu. Type the number to use the following formula.
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
        airTemp_line_begin = int(input("\nWhat's the line on the Excel file that contains the first Air Temperature information?\n Example: If the first air temperature is at the line 13 on the excel document, you must use the value of 12. (1->13 = 12)\n"))
        airTemp_line_end = int(input("\nWhat's the line on the Excel file that contains the last Air Temperature information?\n Example: If the last air temperature is at the line 25 on the excel document, you must use the value of 24. (1->25 = 24)\n"))
        airTemp_col = int(input("\nWhat's the column on the Excel file that contains the Air Temperature information?\n Example: If the air temperature is at the column O on the excel document, you must use the value of 14. (1->15(O) = 14)\n"))
        airTemp_line_aux = airTemp_line_begin
        airTemp_list = []

        dewPon_line_begin = int(input("\nWhat's the line on the Excel file that contains the first Dew Point information?\n Example: If the first dew point is at the line 13 on the excel document, you must use the value of 12. (1->13 = 12)\n"))
        dewPon_line_end = int(input("\nWhat's the line on the Excel file that contains the last Dew Point information?\n Example: If the last dew point is at the line 25 on the excel document, you must use the value of 24. (1->25 = 24)\n"))
        dewPon_col = int(input("\nWhat's the column on the Excel file that contains the Dew Point information?\n Example: If the air temperature is at the column BK on the excel document, you must use the value of 62. (1->63(BK) = 62)\n"))
        dewPon_line_aux = dewPon_line_begin
        dewPon_list = []

        while (airTemp_line_end - airTemp_line_aux >= 0):
            airTemp_list.append(worksheet.cell(airTemp_line_aux, airTemp_col).value)
            airTemp_line_aux = airTemp_line_aux + 1

        while (dewPon_line_end - dewPon_line_aux >= 0) :
            dewPon_list.append(worksheet.cell(dewPon_line_aux, dewPon_col).value)
            dewPon_line_aux = dewPon_line_aux + 1

        print(airTemp_list)
        print(dewPon_list)

        countAir = airTemp_line_end - airTemp_line_begin
        countDew = dewPon_line_end - dewPon_line_begin
        if (countAir != countDew):
            print("The number of information related to Air Temperature is not the same as the number of information related to Dew Point")
        telicynIndex(airTemp_list, dewPon_list, countAir)

    elif option == '3':
        # calculates nesterov index
        print("something")


    elif option == '4':
        relHum_line_begin = int(input("\nWhat's the line on the Excel file that contains the first Relative Humidity information?\n Example: If the first Relative Humidity is at the line 13 on the excel document, you must use the value of 12. (1->13 = 12)\n"))
        relHum_line_end = int(input("\nWhat's the line on the Excel file that contains the last Relative Humidity information?\n Example: If the last Relative Humidity is at the line 25 on the excel document, you must use the value of 24. (1->25 = 24)\n"))
        relHum_col = int(input("\nWhat's the column on the Excel file that contains the Relative Humidity information?\n Example: If the Relative Humidity is at the column O on the excel document, you must use the value of 14. (1->15(O) = 14)\n"))
        relHum_line_aux = relHum_line_begin
        relHum_list = []

        while (relHum_line_end - relHum_line_aux >= 0):
            relHum_list.append(worksheet.cell(relHum_line_aux, relHum_col).value)
            relHum_line_aux = relHum_line_aux + 1
        print(relHum_list)
        monteAlegre(relHum_list, relHum_line_end - relHum_line_begin)

    elif option == '5':
        menu = False

    else:
        print("The input value is invalid. The program is finishing.")
        menu = False
