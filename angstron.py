import xlrd
import math
def angstronIndex (airTemperature, relactiveHumidity):
    ###This function calculates the Angstron Index for a given air temperature and relactive humidity, , measured at 12h.
    result = 0.05 * relactiveHumidity - 0.1 * (airTemperature - 27)
    print('O indice de Angstron do dia eh: ', round(result, 2))


def telicynIndex(airTemperature, dewPoint, N):
    ###This function calculates the Angstron Index for a set of air temperature and dew point, measured at 13h.
    i = 0
    result = 0.0
    while i < N:
        result = math.log(airTemperature - dewPoint, 10) + result
        i = i + 1
    print('O indice de Telicyn do dia e: ', round(result, 2))

def nesterofIndex(airTemperature, airDeficit, N):
    ###This function calculates the Angstron Index for a set of air temperature and air deficit, measured at 13h.
    i = 0
    result = 0.0
    while i < N:
        result = airDeficit - airTemperature + result

    print('O indice de Nesterof e: ', result)

def monteAlegre(relactiveHumidity, N):
    i = 0
    result = 0.0
    while i < N:
        result = (100 / relactiveHumidity) + result
        i = i + 1
    print('Monte Alegre Formula is: ', round(result,2))

workbook = xlrd.open_workbook('/home/thiago/Desktop/Meteorologia/Documents/__RJ_A601_ECOLOGIA_AGRICOLA.xls')
worksheet = workbook.sheet_by_index(0)

airTemp12 = worksheet.cell(12, 13).value
umiRel = worksheet.cell(12, 37).value

airTemp13 = worksheet.cell(12, 14).value
dewPt = worksheet.cell(12, 62).value

umiRel13 = worksheet.cell(12, 38).value

angstronIndex(airTemp12, umiRel)
telicynIndex(airTemp13, dewPt, 1)
#nesterofIndex(airTemp13, )
monteAlegre(umiRel13, 1)
