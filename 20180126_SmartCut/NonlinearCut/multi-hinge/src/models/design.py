"""
read multiple rebar
"""
import pandas as pd

from src.utils.rebar import get_diameter, get_area


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
            for column in list(df):
                if '主筋' in column or '主筋長度' in column or '腰筋' in column:
                    del df[column]
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

        return round(row1 + row2, 8)

    def get_length_area(self, index, abs_length):
        """
        get absolute length correspond rebar area
        """
        index = index // 4 * 4

        left_boundary = (
            self.get(index, ('主筋長度', '左')) + self.get(index, ('支承寬', '左'))
        ) / 100
        right_boundary = (
            self.get(index, ('梁長', '')) -
            self.get(index, ('主筋長度', '右')) +
            self.get(index, ('支承寬', '右'))
        ) / 100

        if abs_length < left_boundary:
            col = ('主筋', '左')
        elif abs_length > right_boundary:
            col = ('主筋', '右')
        else:
            col = ('主筋', '中')

        top = self.get_total_area(index, col)
        bot = self.get_total_area(index + 2, col)

        return top, bot

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

        return float(stirrup.split('@')[1]) / 100

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

    print(design.get_len())
    print(design.get(1))
    print(design.get(3, ('主筋', '左')))
    print(design.get(2, ('主筋長度', '左')))
    print(design.get_total_area(11, ('主筋', '左')))
    print(design.get_length_area(11, 0.24))
    print(design.get_num(6, ('主筋', '右')))
    print(design.get_diameter(6, ('箍筋', '右')))
    print(design.get_diameter(6, ('主筋', '右')))
    print(design.get_area(6, ('主筋', '右')))
    print(design.get_area(6, ('箍筋', '右')))
    print(design.get_spacing(10, ('箍筋', '右')))
    print(design.get_shear(10, ('箍筋', '右')))
    print(design.get_shear(10, ('箍筋', '中')))


if __name__ == "__main__":
    main()
