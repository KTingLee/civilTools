import numpy as np
from scipy import stats
# from matplotlib import pyplot as plt

# diff找delta(網格長寬)
def findDeltaValue(axisCoords):
    delta = np.diff(axisCoords)
    delta = delta[delta > 0]  # 去除0
    return stats.mode(delta, keepdims=True)[0][0]  # 找眾數

def getMaxMinValue(array):
    maxValue = np.amax(array)
    minValue = np.amin(array)
    return maxValue, minValue

# 取得最大最小值，並計算出網格長度
def getGridInfo(xCoords, yCoords):
    maxX, minX = getMaxMinValue(xCoords)
    maxY, minY = getMaxMinValue(yCoords)
    return {
        'deltaX': findDeltaValue(xCoords),
        'deltaY': findDeltaValue(yCoords),
        'maxX': maxX,
        'minX': minX,
        'maxY': maxY,
        'minY': minY
    }

# 初始化完整邊界陣列(利用xy座標算出網格大小後，產生對應範圍的nan陣列)
def genNaNElevationArr(gridInfo):
    # 計算邊界元素數量
    colSize = int((gridInfo['maxX'] - gridInfo['minX']) / gridInfo['deltaX'] + 1)
    rowSize = int((gridInfo['maxY'] - gridInfo['minY']) / gridInfo['deltaY'] + 1)
    return np.empty((rowSize, colSize,)) * np.nan

# 將原始高程資料依照其座標放置到nan array中對應位置
def mappingElevationToNaNArr(elevationData, nanElevationArr, gridInfo):
    # 將原始資料由 x, y 轉換成 indexX, indexY
    for [z, x, y] in elevationData:
        col = (x - gridInfo['minX']) / gridInfo['deltaX']
        row = (gridInfo['maxY'] - y) / gridInfo['deltaY']  # 因為y是由下往上增加，所以計算index的方式要顛倒
        if col.is_integer() and row.is_integer():  # 過濾奇點
            nanElevationArr[int(row)][int(col)] = z
    return nanElevationArr

# 將方格高程資料轉換成四點法資料
def createGridRowsData(gridElevationArr):
    result = []
    rows, cols = gridElevationArr.shape
    for row in range(rows - 1):
        for col in range(cols - 1):
            a = gridElevationArr[row][col]
            b = gridElevationArr[row][col + 1]
            c = gridElevationArr[row + 1][col]
            d = gridElevationArr[row + 1][col + 1]
            elevations = [a, b, c, d]
            if np.isnan(elevations).any():
                continue
            result.append(elevations)
    return result

path = './sourceData/'
fileName = 'B區點位.csv'

# 讀取檔案，讀取後的欄位分別是Z,X,Y
elevationData = np.genfromtxt(f"{path}{fileName}", delimiter=',', skip_header = 1, usecols = (2, 3, 4))

# 排序
elevationData = elevationData[np.lexsort(
    (elevationData[:, 1], elevationData[:, 2])
)]  # 先依y排，再依x排

zCoords, xCoords, yCoords = elevationData.T
gridInfo = getGridInfo(xCoords, yCoords)
nanElevationArr = genNaNElevationArr(gridInfo)

elevationArr = mappingElevationToNaNArr(elevationData, nanElevationArr, gridInfo)
gridElevationRows = createGridRowsData(elevationArr)

np.savetxt('test.csv', gridElevationRows, delimiter=',', fmt='%.3f')
