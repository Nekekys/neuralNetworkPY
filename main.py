import keras
import numpy as np
import openpyxl
import math
from random import randint

table = openpyxl.open("big.xlsx", read_only=True)
sheet = table.active

inputValue = []
for row in sheet.iter_rows(min_col=30, max_col=57, min_row=2):
    rowElem = []
    for cell in row:
        rowElem.append(float(cell.value))
    inputValue.append(rowElem)

inputValueValidate = []
outputValueValidate = []


for row in inputValue:
    rowElemMainI = []
    rowElemMainO = []
    k = 0
    ran = randint(0, 27)
    for item in row:
        item = round(item / 100.0, 4)

        if k == ran:
            rowElemMainO.append(item)
        else:
            rowElemMainI.append(item)

        k += 1

    inputValueValidate.append(rowElemMainI)
    outputValueValidate.append(rowElemMainO)



inputValueValidate = np.array(inputValueValidate)
outputValueValidate = np.array(outputValueValidate)

inputValueValidate_train = inputValueValidate[:200]
outputValueValidate_train = outputValueValidate[:200]

inputValueValidate_check = inputValueValidate[200:]
outputValueValidate_check = outputValueValidate[200:]


model = keras.Sequential()
model.add(keras.layers.Dense(units=27, activation="relu"))
model.add(keras.layers.Dense(units=1, activation="linear"))
model.compile(loss="mse", optimizer="NAdam", metrics=["accuracy"])

# Adadelta - loss 1.3  // error sum - 17.4 / average error - 0.7
# Adagrad - loss 9.3  // error sum - 13.1 / average error - 0.52
# Adam - loss 7 // error sum - 10.2 / average error - 0.4
# Adamax - loss 3.4 // error sum - 8.9 / average error - 0.35
# FTRL - loss 6  // error sum - 10.9 / average error - 0.43
# NAdam - loss 2.7 // error sum - 8.6 / average error - 0.34
# sgd - loss 7.6 // error sum - 9.8 / average error - 0.4
# RMSprop - loss 4 // error sum - 11.3 / average error - 0.45

result = model.fit(inputValueValidate_train, outputValueValidate_train, epochs=10000, validation_split=0.2)

predicted_test = model.predict(inputValueValidate_check)

totalErrorCount = 0
print("Result: ")
for i in range(len(predicted_test)):
    totalErrorCount += math.fabs(predicted_test[i] * 100 - outputValueValidate_check[i] * 100)
    print("neural value ", predicted_test[i] * 100, " ~ ", outputValueValidate_check[i] * 100, " true value", "  (", math.fabs(predicted_test[i] * 100 - outputValueValidate_check[i] * 100) ,")")
print("total error - ", totalErrorCount)
model.save("save_model")

