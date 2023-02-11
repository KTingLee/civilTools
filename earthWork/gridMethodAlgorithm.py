import numpy as np
from scipy import stats
# from matplotlib import pyplot as plt


def findDeltaValue(deltaArray):
    deltaArray = deltaArray[deltaArray > 0]  # 去除0
    return stats.mode(deltaArray, keepdims=True)[0][0]  # 找眾數


path = './sourceData/elevationData.txt'
elevationData = np.loadtxt(path, skiprows=1)

# 排序
elevationData = elevationData[np.lexsort(
    (elevationData[:, 1], elevationData[:, 2])
)]  # 先依y排，再依x排

# diff找delta
delta = np.diff(elevationData[:, [1, 2]], axis=0)  # axis=0: 依row做diff
deltaX = findDeltaValue(delta[:, 0])
deltaY = findDeltaValue(delta[:, 1])

# 找 max, min 做邊界
maxZ, maxX, maxY = np.amax(elevationData, axis=0)  # axis=0: 依column找
minZ, minX, minY = np.amin(elevationData, axis=0)

# 計算邊界元素數量
lengthX = int((maxX - minX) / deltaX + 1)
lengthY = int((maxY - minY) / deltaY + 1)

# 初始化完整邊界陣列
elevationArr = np.empty((lengthY, lengthX,)) * np.nan

# 將原始資料由 x, y 轉換成 indexX, indexY
for [z, x, y] in elevationData:
    col = (x - minX) / deltaX
    row = (maxY - y) / deltaY  # 因為y是由下往上增加，所以計算index的方式要顛倒
    if ((col % 1 > 0.0) | (row % 1 > 0.0)):  # 過濾奇點
        continue
    elevationArr[int(row)][int(col)] = z

result = []
for row in range(elevationArr.shape[0] - 1):
    for col in range(elevationArr.shape[1] - 1):
        a = elevationArr[row][col]
        b = elevationArr[row][col + 1]
        c = elevationArr[row + 1][col]
        d = elevationArr[row + 1][col + 1]
        elevations = [a, b, c, d]
        if np.isnan(elevations).any():
            continue
        result.append(elevations)

np.savetxt('result.csv', result, delimiter=',', fmt='%.3f')
