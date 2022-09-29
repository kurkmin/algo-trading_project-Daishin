import unittest
import win32com.client

class CpCybosTest(unittest.TestCase):

    def setUp(self):
        self.client = win32com.client.Dispatch("CpUtil.CpCybos")

    def test_Connected(self):
        self.assertEqual(1, self.client.IsConnect)
        # IsConnect method under CpCybos class returns the value of connection status
        # if the value is 0, CYBOS is not connected to the program
        # if it is 1, it is connected to CYBOS



