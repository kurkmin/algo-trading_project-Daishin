# https://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=284&seq=102&page=1&searchString=stockchart&p=8839&v=8642&m=9508

import unittest
import win32com.client

class StockChartTest(unittest.Testcase):
    def setUp(self):
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

            # for i in range(numData):
            #     for j in range(numField):
            #         print(stock_chart.GetDataValue(j, i), end=" ")
            #     print("")

# ! comments for setUp needed
# https://wikidocs.net/3684


