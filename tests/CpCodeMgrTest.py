import unittest
import win32com.client as client

class CpCodeMgrTest(unittest.TestCase):
    def setUp(self):
        self.CpCode = client.Dispatch("CpUtil.CpCodeMgr")
    # CpCodeMgr class contains methods for getting information about stock codes

    # def test_