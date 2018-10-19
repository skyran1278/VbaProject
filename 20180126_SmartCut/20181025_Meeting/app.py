import os
import re
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from dataset_beam_design import load_beam_design
from dataset_e2k import load_e2k


rebar = ['#4', '2#4', '2#5', '2#6']
max_spacing = 0.25
min_spacing = 0.10

beam_3points_table = pd.DataFrame(columns=['樓層', '編號', 'RC 梁寬', 'RC 梁深', '主筋', '主筋, 左', '主筋, 中', '主筋, 右',
                                           '長度, 左', '長度, 中', '長度, 右', '腰筋', '箍筋, 左', '箍筋, 中', '箍筋, 右', '梁長', '支承寬', 'NOTE', 'MESSAGE'])

rebars, stories, point_coordinates, lines, materials, sections = load_e2k()

beam_design_table = load_beam_design()

# first time calc
beam_design_table = beam_design_table.assign(dbt=rebar[0],
                                             spacing=lambda x: rebars[rebar[0], 'AREA'] * 2 / x.VRebar)

print(beam_design_table.head())

for name, group in beam_design_table.groupby(['Story', 'BayID']):
    i = 1
    while np.any(group['spacing'] < min_spacing):
        rebar_num, rebar_size = rebar[i].split(sep='#')
        rebar_size = '#' + rebar_size

        if rebar_num == '2':
            def f(x): return rebars[rebar_size, 'AREA'] * 4 / x.VRebar
        else:
            def f(x): return rebars[rebar_size, 'AREA'] * 2 / x.VRebar

        group = group.assign(dbt=rebar[i], spacing=f)

        i += 1


beam_design_table = beam_design_table.assign(
    spacing=lambda x: np.minimum(x.spacing, max_spacing))

# 支承寬要怎麼走
for name, group in beam_design_table.groupby(['Story', 'BayID']):
    i = 1
    group_max = np.amax(group['StnLoc'])
    group_min = np.amin(group['StnLoc'])

    group_left = (group_max - group_min) / 4 + group_min
    group_right = 3 * (group_max - group_min) / 4 + group_min
    group_size = group['dbt'].iloc[0] + '@'

    group_left_spacing = np.amax(
        group['spacing'][group['StnLoc'] <= group_left]) * 100
    group_mid_spacing = np.amax(
        group['spacing'][(group['StnLoc'] >= group_left) & (group['StnLoc'] <= group_right)]) * 100
    group_right_spacing = np.amax(
        group['spacing'][group['StnLoc'] >= group_right]) * 100

    beam_3points_table['箍筋, 左'].loc[i] = group_size + str(group_left_spacing)
    beam_3points_table['箍筋, 中'].loc[i] = group_size + str(group_mid_spacing)
    beam_3points_table['箍筋, 右'].loc[i] = group_size + str(group_right_spacing)

    # print([group['StnLoc'] <= group_left])
    # print(np.amax(group['spacing'][group['StnLoc'] <= group_left]))
    i += 1

print(beam_3points_table.head())
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
plt.figure()
plt.scatter(point_coordinates['X'], point_coordinates['Y'], marker='.')
plt.show()


print('Done')
