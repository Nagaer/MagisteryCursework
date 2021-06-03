import numpy as np

from DataExtractor import Extractor

def maxMin(x, y):
    z = []
    for x1 in x:
        for y1 in y.T:
            z.append(max(np.minimum(x1, y1)))
    return np.array(z).reshape((x.shape[0], y.shape[1]))


def fuzzification(temp):
    N = 0.0
    AN = 0.0
    H = 0.0
    Cr = 0.0
    if temp < 36.6:
        return None
    elif temp < 37.0:
        N = (37.0 - temp)/0.4
        AN = (temp - 36.6)/0.4
    elif temp < 38.0:
        AN = 1.0
    elif temp < 39.0:
        AN = 39.0 - temp
        H = temp - 38.0
    elif temp < 39.5:
        H = 1.0
    elif temp < 40.5:
        H = 40.5 - temp
        Cr = temp - 39.5
    elif temp < 42.0:
        Cr = 1.0
    else:
        return None

    return [round(N, 4), round(AN, 4), round(H, 4), round(Cr, 4)]


A = np.array([[1.0, 0.3, 0, 0.2],
              [1.0, 0, 0, 0.2],
              [0.8, 0.2, 0.1, 0.2],
              [1.0, 0, 0, 0.4],
              [0.7, 0.1, 0.1, 0.2],
              [0.9, 0.3, 0, 0.3],
              [0.7, 0.6, 0, 0.3],
              [1.0, 0, 0, 0.5],
              [0.3, 0.9, 0, 0.3],
              [0.3, 0.8, 0, 0.3],
              [0.2, 0.8, 0, 0.3],
              [0.5, 1.0, 0, 0.2],
              [0.2, 1.0, 0, 0.1],
              [0.1, 1.0, 0, 0.1],
              [0.1, 0, 1.0, 0.4],
              [0.1, 0, 1.0, 0.4],
              [0.1, 0, 1.0, 0.5],
              [0.1, 0, 1.0, 0.5],
              [0.1, 0, 0.2, 1.0],
              [0.1, 0, 0.2, 1.0],
              [0.1, 0, 0.2, 1.0],
              [0.1, 0, 0.2, 1.0],
              [0.1, 0, 0.2, 1.0]])

extractor = Extractor()
risksTemps = extractor.extract(5)
B = np.array([fuzzification(risksTemps[0]), fuzzification(risksTemps[1]),
              fuzzification(risksTemps[2]), fuzzification(risksTemps[3])])

print(maxMin(A, B))
