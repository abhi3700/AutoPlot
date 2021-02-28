import json

# ============================================================================================
# Read from file
with open('./info.json') as f:
    data = json.load(f)

# ============================================================================================
# print all data
# print(data)

# ============================================================================================
# print CP wafer name
print(f'CP wafer_type: {data["wafer_types"]["cp"]}')

# ============================================================================================
# print measurement sites of Nit wafer
coordinates_nit = list(zip(data["measurement_sites"]["Nit"]["x"], data["measurement_sites"]["Nit"]["y"]))
print(f'Nit coordinates: {coordinates_nit}')

# print the x-coordinates of Nit
print(f'Nit\'s x-coordinates: {data["measurement_sites"]["Nit"]["x"]}')
# print the y-coordinates of Nit
print(f'Nit\'s y-coordinates: {data["measurement_sites"]["Nit"]["y"]}')

# print the 1st coordinate of Nit
print(f'Nit\'s 1st coordinate: {coordinates_nit[0]}')
# print the 1st x-coordinate of Nit
print(f'Nit\'s 1st x-coordinate: {coordinates_nit[0][0]}')
# print the 1st y-coordinate of Nit
print(f'Nit\'s 1st y-coordinate: {coordinates_nit[0][1]}')

# --------------------------------------------------------------------------------------------
# print measurement sites of Poly wafer
coordinates_poly = list(zip(data["measurement_sites"]["Poly"]["x"], data["measurement_sites"]["Poly"]["y"]))
print(f'Poly coordinates: {coordinates_poly}')

# print the x-coordinates of Poly
print(f'Poly\'s x-coordinates: {data["measurement_sites"]["Poly"]["x"]}')
# print the y-coordinates of Poly
print(f'Poly\'s y-coordinates: {data["measurement_sites"]["Poly"]["y"]}')

# print the 1st coordinate of Poly
print(f'Poly\'s 1st coordinate: {coordinates_poly[0]}')
# print the 1st x-coordinate of Poly
print(f'Poly\'s 1st x-coordinate: {coordinates_poly[0][0]}')
# print the 1st y-coordinate of Poly
print(f'Poly\'s 1st y-coordinate: {coordinates_poly[0][1]}')

# ============================================================================================
# Read from file
with open('./info.json') as f:
    data = json.load(f)
