import unittest
import pandas as pd

df = pd.read_excel('sample.xlsx')

print(df)


class MyTestCase(unittest.TestCase):
    def test_something(self):
        print('hello world')


if __name__ == '__main__':
    unittest.main()
