"""
read multiple rebar
"""
import pandas as pd

from src.utils.rebar import get_diameter, get_area


# def get_design(path):
#     """
#     get excel
#     """
#     df = pd.read_excel(
#         path, sheet_name='多點斷筋', header=[0, 1], usecols=23)

#     df = df.rename(columns=lambda x: x if 'Unnamed' not in str(x) else '')

#     return df


class Design:
    """
    excel beam design
    """

    def __init__(self, path):
        df = pd.read_excel(
            path, sheet_name='多點斷筋', header=[0, 1], usecols=22)

        df = df.rename(columns=lambda x: x if 'Unnamed' not in str(x) else '')

        self.df = df

    def get_len(self):
        """
        get index length
        """
        return len(self.df.index)

    def get(self, index, column):
        # for others to normalize to first row
        if '主筋' not in column:
            index = index // 4 * 4

        # for 主筋長度 to get first and last row
        elif '主筋長度' in column:
            if index % 4 == 1:
                index -= 1
            if index % 4 == 2:
                index += 1

        # for 主筋 to get its row
        return self.df.loc[index, column]

    # def get_story(self, index):
    #     """
    #     get story
    #     """
    #     index = index // 4 * 4
    #     return self.df.loc[index, ('樓層', '')]

    # def get_id(self, index):
    #     """
    #     get bay id
    #     """
    #     index = index // 4 * 4
    #     return self.df.loc[index, ('編號', '')]

    # def get_span(self, index):
    #     index = index // 4 * 4
    #     return self.df.loc[index, ('梁長', '')]

    def get_num(self, index, column):
        num_and_size = self.get(index, column)
        return int(num_and_size.split('-')[0])

    def get_diameter(self, index, column):
        size = self.get(index, column)

        # 主筋
        if '-' in size:
            size = size.split('-')[1]

        # 箍筋
        elif '@' in size:
            size = size.split('@')[0]

        return get_diameter(size)

    def get_area(self, index, column):
        size = self.get(index, column)

        # 主筋
        if '-' in size:
            size = size.split('-')[1]

        # 箍筋
        elif '@' in size:
            size = size.split('@')[0]

        return get_area(size)

    def get_spacing(self, index, column):
        stirrup = self.get(index, column)

        if '@' not in stirrup:
            raise Exception("Invalid index!", (index, column))

        return float(stirrup.split('@')[1])

    def get_shear(self, index, column):
        area = self.get_area(index, column)
        spacing = self.get_spacing(index, column)
        return area / spacing


def main():
    """
    test
    """
    # pylint: disable=line-too-long
    path = 'D:/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190323 203316 SmartCut.xlsx'
    # path = '/Users/skyran/Documents/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190323 203316 SmartCut.xlsx'

    design = Design(path)

    print(design.df[('樓層', '')])
    print(design.df[('主筋', '左')])
    print(design.df.head())
    # print(design.get_story(5))

    # design = get_design(path)

    # print(design[('樓層', '')])
    # print(design[('主筋', '左')])
    # print(design.head())


if __name__ == "__main__":
    main()
