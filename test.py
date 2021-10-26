import keras
import numpy as np
import openpyxl
import math
import pandas as pd

def arrayEntry(row, arr):
    for elem in arr:
        if row == elem:
            return False
    return True

table = openpyxl.open("big.xlsx", read_only=True)
sheet = table.active


inputValue = []
mediumArray = []
for row in sheet.iter_rows(min_col=30, max_col=57, min_row=2):
    rowElem = []
    sumElem = 0
    for cell in row:
        rowElem.append(float(cell.value))
        sumElem += round(float(cell.value) / 100.0, 4)
    mediumArray.append(sumElem / 28)
    inputValue.append(rowElem)



# df = pd.read_excel("big.xlsx")
# df.head()
# #df.at[1,0]=100
# print(df['A':'B'])
# df.to_excel("testBook.xlsx")

inputData = []
outputData = []

coordinatesDots = []
j = 0
for row in inputValue:
    error = 0
    k = 0
    for cell in row:
        item = round(float(cell)/100, 4)

        if k % 2 == 1:
            if k == 1:
                if math.fabs(round(row[27] / 100.0, 4) - item) >= 0.03 and math.fabs(round(row[3] / 100.0, 4) - item) >= 0.03:
                    error += 1
                    coordinatesDots.append([j, k])
            elif k == 27:
                if math.fabs(round(row[25] / 100.0, 4) - item) >= 0.03 and math.fabs(round(row[1] / 100.0, 4) - item) >= 0.03:
                    coordinatesDots.append([j, k])
                    error += 1
            elif math.fabs(round(row[k+2] / 100.0, 4) - item) >= 0.03 and math.fabs(round(row[k-2] / 100.0, 4) - item) >= 0.03:
                coordinatesDots.append([j, k])
                error += 1
        else:
            if k == 0:
                if math.fabs(round(row[26] / 100.0, 4) - item) >= 0.03 and math.fabs(round(row[2] / 100.0, 4) - item) >= 0.03:
                    coordinatesDots.append([j, k])
                    error += 1
            elif k == 26:
                if math.fabs(round(row[24] / 100.0, 4) - item) >= 0.03 and math.fabs(round(row[0] / 100.0, 4) - item) >= 0.03:
                    coordinatesDots.append([j, k])
                    error += 1
            elif math.fabs(round(row[k+2] / 100.0, 4) - item) >= 0.03 and math.fabs(round(row[k-2] / 100.0, 4) - item) >= 0.03:
                coordinatesDots.append([j, k])
                error += 1
            elif math.fabs(mediumArray[j] - item) >= 0.02:
                coordinatesDots.append([j, k])
                error += 1
        k += 1
    j += 1

new_model = keras.models.load_model("save_model")

endValueArray = []
for item in coordinatesDots:
    rowArr = []
    k = 0
    for elem in inputValue[item[0]]:
        if not k == item[1]:
            rowArr.append(elem)
        k += 1
    emptyArr = []
    emptyArr.append(rowArr)
    test2 = new_model.predict(emptyArr)
    endValueArray.append([item[0], item[1], test2[0][0]])


for elem in endValueArray:
    print("row - ", elem[0] + 2, " column - ", elem[1] + 1, " value - ", elem[2], "  origin - ", inputValue[elem[0]][elem[1]], "  medium - ", mediumArray[elem[0]]*100)


print(mediumArray)
print(len(endValueArray))


# print(test2)
