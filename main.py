import keras
import numpy as np
import openpyxl
import math


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

for row in inputValue:
    rowElemI = []
    roeElemO = []
    check = False
    k = 0
    for item in row:
        item = round(item / 100.0, 4)
        if item > 50 and item < 20:
            check = True
        if k == 0:
            lastItem = item
        elif math.fabs(lastItem - item) > 7:
            check = True
        if k == 27:
            roeElemO.append(item)
        else:
            rowElemI.append(item)
        lastItem = item
        k += 1
    if not check:
        inputValueValidate.append(rowElemI)
        outputValueValidate.append(roeElemO)

# print(inputValueValidate)
# print(outputValueValidate)
print(inputValueValidate)
inputValueValidate = np.array(inputValueValidate)
outputValueValidate = np.array(outputValueValidate)

inputValueValidate_x = inputValueValidate[:180]
outputValueValidate_x = outputValueValidate[:180]

inputValueValidate_y = inputValueValidate[180:]
outputValueValidate_y = outputValueValidate[180:]

model = keras.Sequential()
model.add(keras.layers.Dense(units=27, activation="relu"))
model.add(keras.layers.Dense(units=1, activation="linear"))
model.compile(loss="mse", optimizer="sgd", metrics=["accuracy"])


result = model.fit(inputValueValidate_x, outputValueValidate_x, epochs=10, validation_split=0.2)

predicted_test = model.predict(inputValueValidate_y)

print("Result: ")
for i in range(len(predicted_test)):
    print("neural value ", predicted_test[i], " ~ ", outputValueValidate_y[i], " true value")

#model.save("save_model")

