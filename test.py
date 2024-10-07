import unittest
import inspect, time
from stocklab.agent.ebest import EBest
class TestEBest(unittest.TestCase):
    def setUp(self):
        self.ebest = EBest("DEMO")
        self.ebest.login()
    def tearDown(self):
        self.ebest.logout()
