import unittest
from invoice_generator.data.data_preparer import _to_numeric

class TestDataPreparer(unittest.TestCase):

    def test_to_numeric(self):
        self.assertEqual(_to_numeric("123"), 123)
        self.assertEqual(_to_numeric("123.45"), 123.45)
        self.assertEqual(_to_numeric("1,234.56"), 1234.56)
        self.assertEqual(_to_numeric("abc"), "abc")
        self.assertEqual(_to_numeric(None), None)
        self.assertEqual(_to_numeric(123), 123)
        self.assertEqual(_to_numeric(123.45), 123.45)

if __name__ == '__main__':
    unittest.main()
