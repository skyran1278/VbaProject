import os
import re
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from dataset_beam_design import load_beam_design
from dataset_e2k import load_e2k

# TODO: 樓層

rebar = ['#4', '2#4', '2#5', '2#6']
spacing = [0.1, 0.12, 0.15, 0.18, 0.2, 0.25]


dataset_dir = os.path.dirname(os.path.abspath(__file__))

rebars, stories, point_coordinates, lines, materials, sections = load_e2k()

beam_design_table = load_beam_design()

# first time calc
beam_design_table = beam_design_table.assign(dbt=rebar[0],
                                             spacing=lambda x: rebars[rebar[0], 'AREA'] * 2 / x.VRebar)

# print(beam_design_table.head())

index = pd.MultiIndex.from_tuples([('樓層', ''), ('編號', ''), ('RC 梁寬', ''), ('RC 梁深', ''), ('主筋', ''), ('主筋', '左'), ('主筋', '中'), ('主筋', '右'), (
    '長度', '左'), ('長度', '中'), ('長度', '右'), ('腰筋', ''), ('箍筋', '左'), ('箍筋', '中'), ('箍筋', '右'), ('梁長', ''), ('支承寬', ''), ('NOTE', ''), ('MESSAGE', '')])

beam_3points_table = pd.DataFrame(np.empty(
    [len(beam_design_table.groupby(['Story', 'BayID'], sort=False)) * 4, 19], dtype='<U16'), columns=index)

# print(beam_3points_table.head())

for (Story, BayID), group in beam_design_table.groupby(['Story', 'BayID'], sort=False):
    i = 1
    while np.any(group['spacing'] < spacing[0]):
        rebar_num, rebar_size = rebar[i].split(sep='#')
        rebar_size = '#' + rebar_size

        if rebar_num == '2':
            def f(x): return rebars[rebar_size, 'AREA'] * 4 / x.VRebar
        else:
            def f(x): return rebars[rebar_size, 'AREA'] * 2 / x.VRebar
        # print(group)
        # print(group.index.tolist())

        # group = group.assign(dbt=rebar[i], spacing=f)
        group = group.update(dbt=rebar[i], spacing=f)

        i += 1

    beam_design_table.loc[group.index.tolist()] = group

beam_design_table = beam_design_table.assign(
    spacing=lambda x: np.minimum(x.spacing, spacing[-1]))

i = 0
for name, group in beam_design_table.groupby(['Story', 'BayID'], sort=False):
    group_max = np.amax(group['StnLoc'])
    group_min = np.amin(group['StnLoc'])

    for j in range(3):
        pass

    group_left = (group_max - group_min) / 4 + group_min
    group_right = 3 * (group_max - group_min) / 4 + group_min

    group_size = group['dbt'].iloc[0] + '@'

    beam_3points_table.loc[i, ('箍筋', '左')] = group_size + str(
        np.amax(group['spacing'][group['StnLoc'] <= group_left]) * 100)
    beam_3points_table.loc[i, ('箍筋', '中')] = group_size + str(np.amax(group['spacing'][(
        group['StnLoc'] >= group_left) & (group['StnLoc'] <= group_right)]) * 100)
    beam_3points_table.loc[i, ('箍筋', '右')] = group_size + str(
        np.amax(group['spacing'][group['StnLoc'] >= group_right]) * 100)

    # print([group['StnLoc'] <= group_left])
    # print(np.amax(group['spacing'][group['StnLoc'] <= group_left]))
    i = i + 4

beam_3points_table.to_excel(dataset_dir + '/3pionts.xlsx')
# while group['spacing']:
#     pass

# beam_design_table.groupby()
# spacing = [beam_design[['Story', 'BayID', 'VRebar']],
#            rebars[rebar_size, 'AREA'] * 2 / beam_design['VRebar']]


# # spacin = beam_design[['BayID', 'VRebar']]
# story_bayID = list(
#     map('_'.join, zip(beam_design['Story'], beam_design['BayID'])))
# u, indices, counts = np.unique(
#     story_bayID, return_index=True, return_counts=True)
# print(spacing1)
# plt.figure()
# plt.scatter(point_coordinates['X'], point_coordinates['Y'], marker='.')
# plt.show()


print('Done')
