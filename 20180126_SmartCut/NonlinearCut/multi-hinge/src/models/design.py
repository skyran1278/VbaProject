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

    def get(self, index, col=None):
        """
        get by index or col
        """
        if col is None:
            index = index // 4 * 4
            df = self.df.loc[index].to_dict()
            for col in list(df):
                if '主筋' in col or '主筋長度' in col or '腰筋' in col:
                    del df[col]
            return df

        # for 主筋 to get its row
        if '主筋' in col:
            return self.df.loc[index, col]

        # for 主筋長度 to get first and last row
        if '主筋長度' in col:
            if index % 4 == 1:
                index -= 1
            elif index % 4 == 2:
                index += 1
            return self.df.loc[index, col]

        # for others to normalize to first row
        index = index // 4 * 4
        return self.df.loc[index, col]

    def get_total_area(self, index, col):
        """
        get total rebar area
        """
        if index % 4 <= 1:
            index = index // 4 * 4
        else:
            index = index // 4 * 4 + 2

        row1 = self.get_num(index, col) * self.get_area(index, col)
        row2 = self.get_num(index + 1, col) * self.get_area(index + 1, col)

        return row1 + row2

    def get_num(self, index, col):
        """
        get '主筋' num
        """
        num_and_size = self.get(index, col)

        if num_and_size == 0:
            return 0

        return int(num_and_size.split('-')[0])

    def get_diameter(self, index, col):
        """
        get diameter
        """
        size = self.get(index, col)

        if size == 0:
            return 0

        # 主筋
        if '-' in size:
            size = size.split('-')[1]

        # 箍筋
        elif '@' in size:
            size = size.split('@')[0]

        return get_diameter(size)

    def get_area(self, index, col):
        """
        get area
        """
        size = self.get(index, col)

        if size == 0:
            return 0

        # 主筋
        if '-' in size:
            size = size.split('-')[1]

        # 箍筋
        elif '@' in size:
            size = size.split('@')[0]

        return get_area(size)

    def get_spacing(self, index, col):
        """
        get spacing
        """
        stirrup = self.get(index, col)

        if '@' not in stirrup:
            raise Exception("Invalid index!", (index, col))

        return float(stirrup.split('@')[1])

    def get_shear(self, index, col):
        """
        get shear design
        """
        area = self.get_area(index, col)
        spacing = self.get_spacing(index, col)
        return area / spacing


def main():
    """
    test
    """
    from tests.config import config

    design = Design(config['design_path'])

    print(design.get(1))
    print(design.get(3, ('主筋', '左')))
    print(design.get(2, ('主筋長度', '左')))
    # print(design.get_story(5))

    # design = get_design(path)

    # print(design[('樓層', '')])
    # print(design[('主筋', '左')])
    # print(design.head())


if __name__ == "__main__":
    main()
