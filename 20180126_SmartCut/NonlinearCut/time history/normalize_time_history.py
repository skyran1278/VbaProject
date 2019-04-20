import shlex

import numpy as np
import pandas as pd


def main():
    """
    test
    """
    import os

    # global
    script_folder = os.path.dirname(os.path.abspath(__file__))

    peernga_folder = script_folder + '/PEERNGARecords_Unscaled'

    normalized_folder = script_folder + '/PEERNGARecords_Normalized'

    time_historys = {
        'RSN125_FRIULI.A_A-TMZ000': 1.737,
        'RSN767_LOMAP_G03000':  1.093,
        'RSN1148_KOCAELI_ARE000':  2.845,
        'RSN1602_DUZCE_BOL000':  0.710,
        'RSN1111_KOBE_NIS090':  1.037,
        'RSN1633_MANJIL_ABBAR--L':  0.935,
        'RSN725_SUPER.B_B-POE270':  0.964,
        'RSN68_SFERN_PEL180':  2.343,
        'RSN960_NORTHR_LOS270':  0.965,
        'RSN1485_CHICHI_TCU045-N':  0.856,
    }

    for time_history in time_historys:
        normalized_data = []

        with open(f'{peernga_folder}/{time_history}.AT2') as f:
            contents = f.readlines()

        for line in contents[4:]:
            normalized_data.extend([
                float(i) * time_historys[time_history] for i in shlex.split(line)
            ])

        with open(f'{normalized_folder}/{time_history}.AT2', mode='w', encoding='big5') as f:
            f.writelines(contents[:4])
            f.write('\n'.join([f'{i:.7E}' for i in normalized_data]))


if __name__ == "__main__":
    main()
