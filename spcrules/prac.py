import numpy as np
import math
import statistics as stat

x = np.random.randint(1, 100, 10)
print(x)
print(stat.mean(x))

var = stat.mean(x)/len(x)
sigma3 = math.sqrt(var)

print(sigma3)