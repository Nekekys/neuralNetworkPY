import keras

inputData = []
new_model = keras.models.load_model("save_model")
test2 = new_model.predict(inputData)
print(test2)
