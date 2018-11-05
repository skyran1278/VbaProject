import os

import pandas as pd
import numpy as np

from dataset.dataset_e2k import load_e2k
from dataset.const import BAR
from utils.pkl import load_pkl

dataset_dir = os.path.dirname(os.path.abspath(__file__))


def calc_ld(beam_v_m):
    # It is used for nominal concrete in case of phi_e=1.0 & phi_t=1.0.
    # Reference:土木401-93
    PI = 3.1415926

    rebars, _, _, _, materials, sections = load_e2k()

    def _ld(df, Loc):
        # Loc = Loc.capitalize()

        bar_size = 'Bar' + Loc + 'Size'
        bar_1st = 'Bar' + Loc + '1st'

        # 延伸長度比較熟悉 cm 操作
        # m => cm
        B = df['SecID'].apply(lambda x: sections[x, 'B']) * 100
        material = df['SecID'].apply(lambda x: sections[x, 'MATERIAL'])
        fc = material.apply(lambda x: materials[x, 'FC']) / 10
        fy = material.apply(lambda x: materials[x, 'FY']) / 10
        fyh = fy
        cover = 0.04 * 100
        db = df[bar_size].apply(lambda x: rebars[x, 'DIA']) * 100
        num = df[bar_1st]
        dh = df['VNoDuSize'].apply(lambda x: rebars[x, 'DIA']) * 100
        spacing = df['SetSpacing'] * 100

        # 5.2.2
        fc[np.sqrt(fc) > 26.5] = 700

        # R5.3.4.1.1
        cc = dh + cover

        # R5.3.4.1.1
        cs = (B - db * num - dh * 2 - cover * 2) / (num - 1) / 2

        # Vertical splitting failure / Horizontal splitting failure
        cb = np.where(cc <= cs, cc, cs) + db / 2

        # R5.3.4.1.2
        ktr = np.where(cc <= cs, 1, 2 / num) * (PI * dh ** 2 / 4) * fyh / 105 / spacing

        # if cs > cc:
        #     # Vertical splitting failure
        #     cb = db / 2 + cc
        #     # R5.3.4.1.2
        #     ktr = (PI * dh ** 2 / 4) * fyh / 105 / spacing
        # else:
        #     # Horizontal splitting failure
        #     cb = db / 2 + cs
        #     # R5.3.4.1.2
        #     ktr = 2 * (PI * dh ** 2 / 4) * fyh / 105 / spacing / num

        # 5.3.4.1
        ld = 0.28 * fy / np.sqrt(fc) * db / np.minimum((cb + ktr) / db, 2.5)

        # 5.3.4.1
        simple_ld = 0.19 * fy / np.sqrt(fc) * db

        # phi_s factor
        ld[db < 2.2] = 0.8 * ld
        simple_ld[db < 2.2] = 0.8 * simple_ld

        # phi_t factor
        if Loc == 'Top':
            ld = 1.3 * ld
            simple_ld = 1.3 * simple_ld

        ld[ld > simple_ld] = simple_ld

        # 5.3.1
        ld[ld < 30] = 30

        return {
            # cm => m
            Loc + 'Ld': ld / 100
        }

    for Loc in BAR.keys():
        beam_v_m = beam_v_m.assign(**_ld(beam_v_m, Loc))

    return beam_v_m


def add_ld(beam_v_m_ld):
    def init_ld(df):
        return {
            bar_num_ld: df[bar_num],
            # bar_1st_ld: df[bar_1st],
            # bar_2nd_ld: df[bar_2nd]
        }
    for Loc in BAR.keys():
        # Loc = Loc.capitalize()

        bar_num = 'Bar' + Loc + 'Num'
        ld = Loc + 'Ld'
        bar_num_ld = bar_num + 'Ld'
        # bar_1st_ld = bar_1st + 'Ld'
        # bar_2nd_ld = bar_2nd + 'Ld'

        beam_v_m_ld = beam_v_m_ld.assign(**init_ld(beam_v_m_ld))

        for name, group in beam_v_m_ld.groupby(['Story', 'BayID'], sort=False):
            for i in range(len(group)):
                stn_loc = group.at[group.index[i], 'StnLoc']
                stn_ld = group.at[group.index[i], ld]
                stn_inter = (group['StnLoc'] >= stn_loc - stn_ld) & (group['StnLoc'] <= stn_loc + stn_ld)
                group.loc[group[stn_inter].index, bar_num_ld] = max(
                    group.at[group.index[i], bar_num], group.at[group.index[i], bar_num_ld])

            beam_v_m_ld.loc[group.index, bar_num_ld] = group[bar_num_ld]
            print(name)

    return beam_v_m_ld


def cut_optimization(beam_v_m_ld_add, beam_3p):
    output_loc = {
        'Top': {
            'START_LOC': 0,
            'TO_2nd': 1
        },
        'Bot': {
            'START_LOC': 3,
            'TO_2nd': -1
        }
    }

    return beam_3p


def main():
    beam_v_m = load_pkl(dataset_dir + '/beam_v_m.pkl')

    beam_v_m_ld = calc_ld(beam_v_m)
    beam_v_m_ld_add = add_ld(beam_v_m_ld)


if __name__ == '__main__':
    main()
