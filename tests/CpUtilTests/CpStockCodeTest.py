import unittest
import win32com.client as client

class CpStockCodeTest(unittest.TestCase):
    def setUp(self):
        self.stocks = client.Dispatch("CpUtil.CpStockCode")

    def test_GetCount(self):
        self.assertEqual(3733, self.stocks.GetCount())
        # GetCount() method under CpStockCode class returns the total number of stocks listed on exchanges in Korea

    def test_GetData_code(self):
        self.assertEqual('A000020', self.stocks.getData(0, 0))
        # GetData(type, index) method returns data of stock at a given index, depending on a given type:
        # if the type is 0, it returns the code of stock at a given index

    def test_GetData_name(self):
        self.assertEqual("동화약품", self.stocks.getData(1, 0))
        # if the type is 1, it returns the name (written in Korean) of stock at a given index

    def test_GetData_full_code(self):
        self.assertEqual('KR7000020008', self.stocks.getData(2, 0))
        # if the type is 2, it returns the full code of stock at a given index

    def test_getData_multiple_names(self):
        actual_stocks = ['동화약품', 'KR모터스', '경방']
        expected_stocks = []
        for i in range(0, 3):
            expected_stocks.append(self.stocks.getData(1,i))
        self.assertEqual(actual_stocks, expected_stocks)

    def test_NameToCode(self):
        self.assertEqual('A035420', self.stocks.NameToCode('NAVER'))
    # this method returns the code of stock, given a name of the stock as a parameter

    


