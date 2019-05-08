"""
normalize
"""
import shlex


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
        'RSN68_SFERN_PEL090': 2.054,
        'RSN125_FRIULI.A_A-TMZ270': 1.462,
        'RSN169_IMPVALL.H_H-DLT262': 1.695,
        'RSN174_IMPVALL.H_H-E11230': 1.000,
        'RSN721_SUPER.B_B-ICC090': 1.067,
        'RSN725_SUPER.B_B-POE360': 1.537,
        'RSN752_LOMAP_CAP000': 1.173,
        'RSN767_LOMAP_G03090': 0.982,
        'RSN848_LANDERS_CLW-TR': 1.027,
        'RSN900_LANDERS_YER270': 0.873,
        'RSN953_NORTHR_MUL279': 0.669,
        'RSN960_NORTHR_LOS000': 1.005,
        'RSN1111_KOBE_NIS000': 0.953,
        'RSN1116_KOBE_SHI000': 1.424,
        'RSN1148_KOCAELI_ARE090': 1.113,
        'RSN1158_KOCAELI_DZC180': 0.758,
        'RSN1244_CHICHI_CHY101-N': 0.408,
        'RSN1485_CHICHI_TCU045-E': 0.891,
        'RSN1602_DUZCE_BOL090': 0.677,
        'RSN1633_MANJIL_ABBAR--T': 0.882,
        'RSN1787_HECTOR_HEC090': 0.996,
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
