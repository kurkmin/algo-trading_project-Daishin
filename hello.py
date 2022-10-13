# codes to be added
import win32com.client

# under StockChartTest
stock_chart = win32com.client.Dispatch("CpSysDib.StockChart")
stock_chart.SetInputValue(0, "A003540")
stock_chart.SetInputValue(1, ord('1'))
stock_chart.SetInputValue(2, 20220501)
stock_chart.SetInputValue(3, 20220101)
# stock_chart.SetInputValue(1, ord('2')) - this is for how many m / w / d
stock_chart.SetInputValue(4, 10)
stock_chart.SetInputValue(5, (0, 2, 3, 4, 5, 8))
stock_chart.SetInputValue(6, ord('D'))
stock_chart.SetInputValue(9, ord('1'))

stock_chart.BlockRequest()

numData = stock_chart.GetHeaderValue(3)
numField = stock_chart.getHeaderValue(1)
for i in range(numData):
    for j in range(numField):
        print(stock_chart.GetDataValue(j, i), end=" ")
    print("")

# under somewhere

import win32com.client

instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
instMarketEye = win32com.client.Dispatch("CpSysDib.MarketEye")

tarketCodeList = instCpCodeMgr.GetGroupCodeList(5)

# Get PER
instMarketEye.SetInputValue(0, 67)
instMarketEye.SetInputValue(1, tarketCodeList)

# BlockRequest
instMarketEye.BlockRequest()

# GetHeaderValue
numStock = instMarketEye.GetHeaderValue(2)

# GetData
sumPer = 0
for i in range(numStock):
    sumPer += instMarketEye.GetDataValue(0, i)

print("Average PER: ", sumPer / numStock)