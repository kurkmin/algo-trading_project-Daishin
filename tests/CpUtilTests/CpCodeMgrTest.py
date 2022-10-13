import unittest
import win32com.client as client

class CpCodeMgrTest(unittest.TestCase):
    def setUp(self):
        self.CpCodeMgr = client.Dispatch("CpUtil.CpCodeMgr")
    # CpCodeMgr class contains methods for getting information about stock codes

    def test_CodeToName(self):
        self.assertEqual("NAVER", self.CpCodeMgr.CodeToName('A035420'))
    # this method returns the name of stock, given its code as a parameter

    def test_GetStockListByMarket(self):
        self.assertEqual(('A000020', 'A000040', 'A000050'), self.CpCodeMgr.GetStockListByMarket(1)[0:3])
    # ! the documentation for this method is poorly written so it should be examined later
    # https://wikidocs.net/3686 states GetStockListByMarket(1) returns all stocks' codes in tuple

    def test_GetStockSectionKind(self):
        expected_type_code = 1
        # In the documentation, stock type binds to integer value 1
        # similarly, ETF type = 10 whereas ETN type = 17
        actual_type_code = self.CpCodeMgr.GetStockSectionKind('A000020')
        # ! this method returns the type, given a code of (종목)
        self.assertEqual(expected_type_code, actual_type_code)

    def test_GetIndustryCode(self):
        expected_value = "종합주가지수"
        industry_code = self.CpCodeMgr.GetIndustryList()[0]
        actual_value = self.CpCodeMgr.GetIndustryName(industry_code)
        self.assertEqual(expected_value, actual_value)

    def test_GetGroupCode(self):
        expected_value = "CJ씨푸드"
        stock_code = self.CpCodeMgr.GetGroupCodeList(5)[0]
        actual_value = self.CpCodeMgr.CodeToName(stock_code)
        self.assertEqual(expected_value, actual_value)




