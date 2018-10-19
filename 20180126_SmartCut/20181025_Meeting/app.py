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
    group_max = np.amax(group['StnLoc'])
    group_min = np.amin(group['StnLoc'])

    group_left = (group_max - group_min) / 4 + group_min
    group_right = 3 * (group_max - group_min) / 4 + group_min

    print([group['StnLoc'] <= group_left])
    print(group['spacing'][group['StnLoc'] <= group_left])
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
