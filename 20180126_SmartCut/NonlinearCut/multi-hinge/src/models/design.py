"""
read multiple rebar
"""
import pandas as pd

from src.utils.rebar import get_diameter, get_area


class Design:
    """
    excel beam design
    """

    def __init__(self, path, sheet_name='多點斷筋'):
        df = pd.read_excel(
            path, sheet_name=sheet_name, header=[0, 1])

        df = df.rename(columns=lambda x: x if 'Unnamed' not in str(x) else '')

        self.df = df
        self.group_num = self.get_group_num()

    def get_len(self):
        """
        get length
        """
        return len(self.df.index), len(self.df.columns)

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

        if isinstance(col, int):
            col = self.df.columns[col]

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

        return round(row1 + row2, 7)

    def get_abs_length(self, index, col):
        """
        get absolute length
        return start length and end length
        """
        group_num = self.group_num

        length = self.get(index, ('支承寬', '左'))

        if not isinstance(col, int):
            raise Exception('col is not int')

        for iter_col in range(5 + group_num, col):
            length = length + self.get(index, iter_col)

        end_length = length + self.get(index, col)

        return length / 100, end_length / 100

    def get_group_num(self):
        """
        get group num
        """
        index = 1
        while ('主筋', f'左{index}') in self.df.columns:
            index += 1

        index -= 1

        index *= 2

        if ('主筋', '中') in self.df.columns:
            index += 1

        return index

    def get_colname_by_length(self, index, abs_length):
        """
        get absolute length correspond column name
        """
        group_num = self.group_num

        rebar_col = None
        stirrup_col = None

        # m => cm, because design is cm
        abs_length = abs_length * 100

        length = self.get(index, ('支承寬', '左'))
        for col in ('左', '中', '右'):
            length = length + self.get(index, ('箍筋長度', col))
            if abs_length < length:
                stirrup_col = col
                break

        length = self.get(index, ('支承寬', '左'))
        for col in range(5 + group_num, 5 + 2 * group_num):
            length = length + self.get(index, col)
            if abs_length < length:
                rebar_col = self.df.columns[col][1]
                break

        if stirrup_col is None:
            rebar_col = self.df.columns[5 + 2 * group_num - 1][1]
            stirrup_col = '右'

        return rebar_col, stirrup_col

    def get_area_by_length(self, index, abs_length):
        """
        get absolute length correspond rebar area
        """
        index = index // 4 * 4

        top_col = self.get_colname_by_length(index, abs_length)[0]
        bot_col = self.get_colname_by_length(index + 2, abs_length)[0]

        top = self.get_total_area(index, ('主筋', top_col))
        bot = self.get_total_area(index + 2, ('主筋', bot_col))

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
    print(design.get(3, ('主筋', '左1')))
    print(design.get(2, ('主筋長度', '左1')))
    print(design.get_total_area(11, ('主筋', '左1')))
    print(design.get_area_by_length(11, 0.24))
    print(design.get_num(6, ('主筋', '右1')))
    print(design.get_diameter(6, ('箍筋', '右')))
    print(design.get_diameter(6, ('主筋', '右1')))
    print(design.get_area(6, ('主筋', '右1')))
    print(design.get_area(6, ('箍筋', '右')))
    print(design.get_spacing(10, ('箍筋', '右')))
    print(design.get_shear(10, ('箍筋', '右')))
    print(design.get_shear(10, ('箍筋', '中')))

    print(design.get_colname_by_length(0, 0.24))


if __name__ == "__main__":
    main()
