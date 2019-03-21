"""
read multiple rebar
"""
import pandas as pd


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
        return len(self.df.index)

    def get(self, index=None):
        if index is None:
            return self.df

        return


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

    design = get_design(path)

    print(design[('樓層', '')])
    print(design[('主筋', '左')])
    print(design.head())


if __name__ == "__main__":
    main()
