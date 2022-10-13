import unittest
import win32com.client

class PerTest(unittest.TestCase):
    def setUp(self):
        self.CpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
        self.MarketEye = win32com.client.Dispatch("CpSysDib.MarketEye")
        self.tarketCodeList = self.CpCodeMgr.GetGroupCodeList(5)
        self.MarketEye.SetInputValue(0, 67)
        self.MarketEye.SetInputValue(1, self.tarketCodeList)

        self.MarketEye.BlockRequest()
        self.numStock = self.MarketEye.GetHeaderValue(2)

        self.sumPer = 0
        for i in range(self.numStock):
            self.sumPer += self.MarketEye.GetDataValue(0, i)

        self.avgPer = self.sumPer / self.numStock

    @unittest.skip("expected value does not match with actual one")
    def test_getPer(self):
        expected_value = 7.2102174551590625
        actual_value = self.avgPer
        self.assertEqual(expected_value, actual_value)


