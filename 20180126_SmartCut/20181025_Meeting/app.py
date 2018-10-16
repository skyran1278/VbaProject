import os
import re
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from dataset_beam_design import load_beam_design
from dataset_e2k import load_e2k

rebars, stories, point_coordinates, lines, materials, sections = load_e2k()

beam_design = load_beam_design()

rebar_size = '#4'
max_spacing = 0.25
min_spacing = 0.10

a = beam_design.groupby(['Story', 'BayID'])

spacing = [beam_design[['Story', 'BayID', 'VRebar']],
           rebars[rebar_size, 'AREA'] * 2 / beam_design['VRebar']]


# spacin = beam_design[['BayID', 'VRebar']]
story_bayID = list(
    map('_'.join, zip(beam_design['Story'], beam_design['BayID'])))
u, indices, counts = np.unique(
    story_bayID, return_index=True, return_counts=True)
# print(spacing1)
plt.figure()
plt.scatter(point_coordinates['X'], point_coordinates['Y'], marker='.')
plt.show()


print('Done')
