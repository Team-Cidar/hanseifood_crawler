import pandas as pd

pd.set_option('max_rows', None)  # print row without collapsing
pd.set_option('display.width', None)  # print columns without collapsing
pd.set_option('max_colwidth', None)  # print every cells without collapsing


class ExcelParser:
    def parse(self, file_path):
        excel: pd.DataFrame = pd.read_excel(file_path, engine='openpyxl')

        menus_with_target: pd.DataFrame = excel.iloc[3:11]  # 대상 포함
        menus_without_target: pd.DataFrame = menus_with_target.iloc[:, 1:]  # 대상 미포함

        menus: pd.DataFrame = menus_without_target.transpose()
        menus.drop(4, axis=1, inplace=True)  # drop useless column

        print(menus.values)
        for daily_menus in menus.values:
            date = daily_menus[0]
            base_menus = daily_menus[1:-1]
            additional_menus = daily_menus[-1].split('\n')

            print(date)
            print(base_menus)
            print(additional_menus)


if __name__ == '__main__':
    parser = ExcelParser()
    parser.parse('./datas/20230902.xlsx')