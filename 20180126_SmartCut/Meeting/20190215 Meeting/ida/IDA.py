"""
plot ida
"""
import os

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import interp1d

from pushover import Pushover
from dataset import dataset_ida_storydrifts
from interp_IDAS import interp_IDAS
from plot_single_IDA import plot_single_IDA
from plot_multi_IDAS import plot_multi_IDAS
from plot_fractiles import plot_fractiles, plot_fractiles_log
from plot_capacity_rule import plot_DM_rule, plot_IM_rule
from plot_normal_versus_multi import plot_normal_versus_multi, plot_normal_versus_multi_log

# 建議 scaled 到差不多的大小，因為會取最小的來做 median。
# TODO: 把圖拆開
stories = {
    'RF': 4,
    '3F': 3,
    '2F': 2,
}

earthquakes = {
    'elcentro': {
        'pga': 0.214,
        'sa': 0.414
    },
    'TAP010': {
        'pga': 0.117,
        'sa': 0.171,
    },
    'TCU052': {
        'pga': 0.447,
        'sa': 0.683
    },
    'TCU067': {
        'pga': 0.498,
        'sa': 1.234
    },
    'TCU068': {
        'pga': 0.511,
        'sa': 1.383
    },
}

loadcases = [
    'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-1', 'PUSHX-2', 'PUSHX-3', 'PUSHX-MMC', 'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
]

file_dir = os.path.dirname(os.path.abspath(__file__))

pushover = Pushover(story_drifts_path=file_dir + '/20190212 pushover story drifts',
                    base_shear_path=file_dir + '/20190212 pushover base shear',
                    stories=stories, loadcases=loadcases)


multi_story_drifts = dataset_ida_storydrifts(
    '20190124 multi story drifts', stories)
# print(multi_story_drifts.head())
# plot_single_IDA('TCU067', earthquakes, multi_story_drifts, ylim_max=3)

# plot_multi_IDAS(earthquakes, multi_story_drifts,
#                 ylim_max=None, xlim_max=None, accel_unit='pga')
plot_multi_IDAS(earthquakes, multi_story_drifts,
                ylim_max=None, xlim_max=0.025, accel_unit='sa')

# plot_fractiles(earthquakes, multi_story_drifts, ylim_max=2, accel_unit='pga')
# plot_fractiles_log(earthquakes, multi_story_drifts,
#                    ylim_max=2, accel_unit='pga')

# plot_fractiles(earthquakes, multi_story_drifts,
#                ylim_max=4, xlim_max=0.025, accel_unit='sa')
# plot_fractiles_log(earthquakes, multi_story_drifts,
#                    ylim_max=4, xlim_max=0.25, accel_unit='sa')

# multi_x, multi_y = interp_IDAS(
#     earthquakes, multi_story_drifts, accel_unit='sa')
# multi_x = multi_x.quantile(
#     0.5, axis=1, interpolation='nearest').values
# plot_DM_rule(multi_x, multi_y, ylim_max=4,
#              xlim_max=0.25, accel_unit='sa', C_DM=0.02)
# plot_IM_rule(multi_x, multi_y, ylim_max=4,
#              xlim_max=0.25, accel_unit='sa', C_IM=2.03)


normal_story_drifts = dataset_ida_storydrifts(
    '20190124 normal story drifts', stories)
# plot_multi_IDAS(earthquakes, normal_story_drifts, ylim_max=2, accel_unit='pga')
# plot_multi_IDAS(earthquakes, normal_story_drifts, ylim_max=4, accel_unit='sa')

# plot_fractiles(earthquakes, normal_story_drifts, ylim_max=4, accel_unit='sa')
# plot_fractiles_log(earthquakes, normal_story_drifts,
#                    ylim_max=4, accel_unit='sa')

# normal_x, normal_y = interp_IDAS(
#     earthquakes, normal_story_drifts, accel_unit='sa')
# normal_x = normal_x.quantile(
#     0.5, axis=1, interpolation='nearest').values
# plot_DM_rule(normal_x, normal_y, ylim_max=4,
#              xlim_max=0.025, accel_unit='sa', C_DM=0.02)
# plot_IM_rule(normal_x, normal_y, ylim_max=4,
#              xlim_max=0.025, accel_unit='sa', C_IM=0.77)

for loadcase in loadcases:
    pushover.plot_in_drift_and_accel(loadcase)
plot_normal_versus_multi(earthquakes, normal_story_drifts,
                         multi_story_drifts, ylim_max=3, xlim_max=0.025, accel_unit='sa')
plot_normal_versus_multi_log(earthquakes, normal_story_drifts,
                             multi_story_drifts, ylim_max=3, xlim_max=0.025, accel_unit='sa')
plot_normal_versus_multi(earthquakes, normal_story_drifts,
                         multi_story_drifts, ylim_max=1.25, xlim_max=0.025, accel_unit='pga')
plot_normal_versus_multi_log(earthquakes, normal_story_drifts,
                             multi_story_drifts, ylim_max=1.25, xlim_max=0.025, accel_unit='pga')

plt.show()
