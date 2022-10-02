# https://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=284&seq=131&page=1&searchString=marketEye&p=8839&v=8642&m=9508

import unittest
import win32com.client


class MarketEyeTest(unittest.TestCase):
    def setUp(self):
        market_eye = win32com.client.Dispatch("CpSysDib.MarketEye")
        market_eye.SetInputValue(0, (4, 67, 70, 111))
        market_eye.SetInputValue(1, 'A003540')
        market_eye.BlockRequest()
        # print("현재가: ", instMarketEye.GetDataValue(0, 0))
        # print("PER: ", instMarketEye.GetDataValue(1, 0))
        # print("EPS: ", instMarketEye.GetDataValue(2, 0))
        # print("최근분기년월: ", instMarketEye.GetDataValue(3, 0))
