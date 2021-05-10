import unittest
import pandas as pd


class MyTestCase(unittest.TestCase):
    def df_io(self, src, res):
        self.df_dlr = pd.read_excel(src, sheet_name='代理人')
        self.df_xz = pd.read_excel(src, sheet_name='险种')

        self.df_dlr.to_csv(res + '.csv', encoding='utf-8')
        writer = pd.ExcelWriter(res + '.xlsx')
        self.df_dlr.to_excel(writer, sheet_name='代理人')
        writer.save()
        writer.close()

    def test_sum_ldr(self):
        self.df_io('resources/20210506.xlsx', 'resources/results/result')

        # =SUMIF(险种!E:E,E:E,险种!R:R)-SUMIFS(险种!R:R,险种!U:U,"终止",险种!E:E,E:E)
        for id_code, series in self.df_xz.groupby('业务员代码'):
            self.df_dlr.loc[
                self.df_dlr['业务员代码'] == id_code,  # 险种表格和代理人表格 id 相同的 row
                '全月预收规保'  # 此 row 中的'全月预收规保' col
            ] = series['预收期交保费'].sum()  # 险种表格关于此 id 的'预收期交保费'的 sum()

        print(self.df_dlr['全月预收规保'])


if __name__ == '__main__':
    unittest.main()
