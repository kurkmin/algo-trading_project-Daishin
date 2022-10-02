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
        actualType = 1
        # In the documentation, stock type binds to integer value 1
        # similarly, ETF type = 10 whereas ETN type = 17
        expectedType = self.CpCodeMgr.GetStockSectionKind('A000020')
        # ! this method returns the type, given a code of (종목)
        self.assertEqual(actualType, expectedType)



