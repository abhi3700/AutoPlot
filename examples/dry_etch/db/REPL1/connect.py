import json

with open("./info.json") as f:
    data = f.read()
    # data = json.load(f)

# print all data
print(data)

# print CP wafer name
# print(json.dumps(data)["wafer_types"])
print(data["wafer_types"][0])