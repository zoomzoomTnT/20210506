import unittest
import pandas as pd


class MyTestCase(unittest.TestCase):
    def test_something(self):
        self.df_dlr = pd.read_excel('resources/20210506.xlsx', sheet_name='代理人')
        self.df_xz = pd.read_excel('resources/20210506.xlsx', sheet_name='险种')

        print(self.df_xz['预收期交保费'][0])

        # =SUMIF(险种!E:E,E:E,险种!R:R)-SUMIFS(险种!R:R,险种!U:U,"终止",险种!E:E,E:E)
        for i in range(len(self.df_dlr['业务员代码'])):
            id_sum = 0  # starts with 0 for each id

            for j in range(len(self.df_xz['业务员代码'])):
                if self.df_xz['业务员代码'][j] == self.df_dlr['业务员代码'][i]:
                    id_sum += self.df_xz['预收期交保费'][j]

            self.df_dlr['全月预收规保'][i] = id_sum + 0.1

        print(self.df_dlr['全月预收规保'])

        writer = pd.ExcelWriter('resources/result.xlsx')
        self.df_dlr.to_excel(writer, sheet_name = '代理人')
        writer.save()
        writer.close()


if __name__ == '__main__':
    unittest.main()
