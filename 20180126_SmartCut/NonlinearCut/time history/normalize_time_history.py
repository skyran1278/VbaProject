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

    # 'RSN125_FRIULI.A_A-TMZ000': 1.737,
    # 'RSN767_LOMAP_G03000':  1.093,
    # 'RSN1148_KOCAELI_ARE000':  2.845,
    # 'RSN1602_DUZCE_BOL000':  0.710,
    # 'RSN1111_KOBE_NIS090':  1.037,
    # 'RSN1633_MANJIL_ABBAR--L':  0.935,
    # 'RSN725_SUPER.B_B-POE270':  0.964,
    # 'RSN68_SFERN_PEL180':  2.343,
    # 'RSN960_NORTHR_LOS270':  0.965,
    # 'RSN1485_CHICHI_TCU045-N':  0.856,

    time_historys = {
        'RSN68_SFERN_PEL090': 1.837656056,
        'RSN125_FRIULI.A_A-TMZ270': 1.307654483,
        'RSN169_IMPVALL.H_H-DLT262': 1.516095929,
        'RSN174_IMPVALL.H_H-E11230': 0.894655393,
        'RSN721_SUPER.B_B-ICC090': 0.954923036,
        'RSN725_SUPER.B_B-POE360': 1.375469811,
        'RSN752_LOMAP_CAP000': 1.049543505,
        'RSN767_LOMAP_G03090': 0.878556955,
        'RSN848_LANDERS_CLW-TR': 0.919187962,
        'RSN900_LANDERS_YER270': 0.7806415,
        'RSN953_NORTHR_MUL279': 0.59820344,
        'RSN960_NORTHR_LOS000': 0.899213273,
        'RSN1111_KOBE_NIS000': 0.852332215,
        'RSN1116_KOBE_SHI000': 1.273993165,
        'RSN1148_KOCAELI_ARE090': 0.99610448,
        'RSN1158_KOCAELI_DZC180': 0.677974744,
        'RSN1244_CHICHI_CHY101-N': 0.365440287,
        'RSN1485_CHICHI_TCU045-E': 0.796875624,
        'RSN1602_DUZCE_BOL090': 0.60577989,
        'RSN1633_MANJIL_ABBAR--T': 0.788885593,
        'RSN1787_HECTOR_HEC090': 0.891316977,
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
            f.writelines([f'{i:.7E}\n' for i in normalized_data])


if __name__ == "__main__":
    main()
