"""
read multiple rebar
"""
import pandas as pd

from src.utils.rebar import get_diameter, get_area


def get_design(path):
    """
    get excel
    """
    df = pd.read_excel(
        path, sheet_name='多點斷筋', header=[0, 1], usecols=19)

    df = df.rename(columns=lambda x: x if 'Unnamed' not in str(x) else '')

    return df


class Design:
    """
    excel beam design
    """

    def __init__(self, path):
        df = pd.read_excel(
            path, sheet_name='多點斷筋', header=[0, 1], usecols=19)

        df = df.rename(columns=lambda x: x if 'Unnamed' not in str(x) else '')

        self.df = df

    def get_len(self):
        """
        get index length
        """
        return len(self.df.index)

    # def get(self, index=None, column=None):
    #     if index is None:
    #         return self.df

    #     if column is None:
    #         return self.df[index]

    #     if '主筋' not in column:
    #         index = index // 4 * 4
    #         return self.df[index, column]

    #     return self.df[index, column]

    def get_story(self, index):
        """
        get story
        """
        index = index // 4 * 4
        return self.df.loc[index, ('樓層', '')]

    def get_id(self, index):
        """
        get bay id
        """
        index = index // 4 * 4
        return self.df.loc[index, ('編號', '')]

    def get_span(self, index):
        index = index // 4 * 4
        return self.df.loc[index, ('梁長', '')]

    def get_num(self, index, column):
        num_and_size = self.df.loc[index, column]
        return int(num_and_size.split('-')[0])

    def get_diameter(self, index, column):
        size = self.df.loc[index, column]

        # 主筋
        if '-' in size:
            size = size.split('-')[1]

        # 箍筋
        elif '@' in size:
            size = size.split('-')[0]

        return get_diameter(size)

    def get_area(self, index, column):
        size = self.df.loc[index, column]

        # 主筋
        if '-' in size:
            size = size.split('-')[1]

        # 箍筋
        elif '@' in size:
            size = size.split('-')[0]

        return get_area(size)


def main():
    """
    test
    """
    # pylint: disable=line-too-long
    path = 'D:/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190312 235751 SmartCut.xlsx'
    # path = '/Users/skyran/Documents/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190312 235751 SmartCut.xlsx'

    design = Design(path)

    print(design.df[('樓層', '')])
    print(design.df[('主筋', '左')])
    print(design.df.head())
    print(design.get_story(5))

    # design = get_design(path)

    # print(design[('樓層', '')])
    # print(design[('主筋', '左')])
    # print(design.head())


if __name__ == "__main__":
    main()
