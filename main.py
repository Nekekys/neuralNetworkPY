import keras
import numpy as np
import openpyxl
import math
from random import randint

table = openpyxl.open("big.xlsx", read_only=True)
sheet = table.active

inputValue = []
for row in sheet.iter_rows(min_col=30, max_col=57, min_row=2, max_row=226):
    rowElem = []
    for cell in row:
        rowElem.append(float(cell.value))
    inputValue.append(rowElem)

inputValueValidate = []
outputValueValidate = []

inputValueExamination = []
outputValueExamination = []

for row in inputValue:
    rowElemMainI = []
    rowElemMainO = []
    rowElemI = []
    rowElemO = []
    check = False
    k = 0
    ran = randint(0, 27)
    for item in row:
        item = round(item / 100.0, 4)

        if k % 2 == 1:
            if k == 1:
                if math.fabs(round(row[27] / 100.0, 4) - item) >= 0.07 and math.fabs(round(row[3] / 100.0, 4) - item) >= 0.07:
                    rowElemI = row.copy()
                    del rowElemI[k]
                    rowElemO.append(row[k])
                    inputValueValidate.append(rowElemI)
                    outputValueValidate.append(rowElemO)
                    inputValueExamination.append(rowElemI)
                    outputValueExamination.append(rowElemO)
                    rowElemI = []
                    rowElemO = []
            elif k == 27:
                if math.fabs(round(row[25] / 100.0, 4) - item) >= 0.07 and math.fabs(round(row[1] / 100.0, 4) - item) >= 0.07:
                    rowElemI = row.copy()
                    del rowElemI[k]
                    rowElemO.append(row[k])
                    inputValueValidate.append(rowElemI)
                    outputValueValidate.append(rowElemO)
                    inputValueExamination.append(rowElemI)
                    outputValueExamination.append(rowElemO)
                    rowElemI = []
                    rowElemO = []
            elif math.fabs(round(row[k+2] / 100.0, 4) - item) >= 0.07 and math.fabs(round(row[k-2] / 100.0, 4) - item) >= 0.07:
                rowElemI = row.copy()
                del rowElemI[k]
                rowElemO.append(row[k])
                inputValueValidate.append(rowElemI)
                outputValueValidate.append(rowElemO)
                inputValueExamination.append(rowElemI)
                outputValueExamination.append(rowElemO)
                rowElemI = []
                rowElemO = []

        else:
            if k == 0:
                if math.fabs(round(row[26] / 100.0, 4) - item) >= 0.07 and math.fabs(round(row[2] / 100.0, 4) - item) >= 0.07:
                    rowElemI = row.copy()
                    del rowElemI[k]
                    rowElemO.append(row[k])
                    inputValueValidate.append(rowElemI)
                    outputValueValidate.append(rowElemO)
                    inputValueExamination.append(rowElemI)
                    outputValueExamination.append(rowElemO)
                    rowElemI = []
                    rowElemO = []

            elif k == 26:
                if math.fabs(round(row[24] / 100.0, 4) - item) >= 0.07 and math.fabs(round(row[0] / 100.0, 4) - item) >= 0.07:
                    rowElemI = row.copy()
                    del rowElemI[k]
                    rowElemO.append(row[k])
                    inputValueValidate.append(rowElemI)
                    outputValueValidate.append(rowElemO)
                    inputValueExamination.append(rowElemI)
                    outputValueExamination.append(rowElemO)
                    rowElemI = []
                    rowElemO = []
            elif math.fabs(round(row[k+2] / 100.0, 4) - item) >= 0.07 and math.fabs(round(row[k-2] / 100.0, 4) - item) >= 0.07:
                rowElemI = row.copy()
                del rowElemI[k]
                rowElemO.append(row[k])
                inputValueValidate.append(rowElemI)
                outputValueValidate.append(rowElemO)
                inputValueExamination.append(rowElemI)
                outputValueExamination.append(rowElemO)
                rowElemI = []
                rowElemO = []


        if k == ran:
            rowElemMainO.append(item)
        else:
            rowElemMainI.append(item)

        k += 1

    inputValueValidate.append(rowElemMainI)
    outputValueValidate.append(rowElemMainO)



inputValueValidate = np.array(inputValueValidate)
outputValueValidate = np.array(outputValueValidate)

inputValueExamination = np.array(inputValueExamination)
outputValueExamination = np.array(outputValueExamination)


model = keras.Sequential()
model.add(keras.layers.Dense(units=27, activation="relu"))
model.add(keras.layers.Dense(units=1, activation="linear"))
model.compile(loss="mse", optimizer="RMSprop", metrics=["accuracy"])

# Adadelta - loss 140
# Adagrad - loss 20
# Adam - loss 12
# Adamax - loss 18
# FTRL - loss 14
# NAdam - loss 12
# sgd - loss 130
# RMSprop - loss 11

result = model.fit(inputValueValidate, outputValueValidate, epochs=10000, validation_split=0.2)

predicted_test = model.predict(inputValueExamination)

print("Result: ")
for i in range(len(predicted_test)):
    print("neural value ", predicted_test[i], " ~ ", outputValueExamination[i], " true value")

model.save("save_model")

