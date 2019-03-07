"""
plot ida
"""
import os

import matplotlib.pyplot as plt
import scipy.io as sio
import numpy as np

from pushover import Pushover
from ida import IDA
from matlab import Matlab

# 建議 scaled 到差不多的大小，因為會取最小的來做 median。


def main():
    """
    test
    """
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

    # loadcases = [
    #     'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-1', 'PUSHX-2', 'PUSHX-3',
    #     'PUSHX-MMC', 'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
    # ]

    # loadcases = [
    #     'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
    # ]

    file_dir = os.path.dirname(os.path.abspath(__file__))

    pushover_path = {
        'base_shear_path': file_dir + '/20190214 multi pushover base shear',
        'story_drifts_path': file_dir + '/20190214 multi pushover story drifts',
        'story_displacements_path': file_dir + '/20190214 multi pushover displacement'
    }

    ida_path = {
        'base_shear_path': file_dir + '/20190220 multi ida base shear',
        'story_drifts_path': file_dir + '/20190214 multi ida story drifts',
        'story_displacements_path': file_dir + '/20190214 multi ida displacement'
    }

    multi_pushover = Pushover(path=pushover_path, stories=stories)

    ida = IDA(
        path=ida_path,
        earthquakes=earthquakes,
        stories=stories
    )

    # ida.figure(xlim_max=0.02, intensity_measure='pga')
    # ida.plot_all()
    # ida.plot_median()
    # plt.legend(loc='upper left')
    matlab = Matlab(file_dir + '/pushover.mat')

    ida.figure(
        xlim_max=200,
        intensity_measure='sa',
        damage_measure='story_displacements'
    )
    # matlab.plot('triangle', label='Pushover Triangle')
    # ida.plot_all()
    ida.plot_median()
    matlab.plot_all()
    # multi_pushover.plot([
    #     'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-MMCUSER'
    # ])
    plt.legend(loc='upper left')

    # ida.figure(xlim_max=0.02, intensity_measure='sa')
    # # ida.plot_all()
    # ida.plot_median()
    # multi_pushover.plot([
    #     'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
    # ])
    # plt.legend(loc='upper left')

    # ida.figure(xlim_max=0.02, intensity_measure='base_shear')
    # # ida.plot('TCU067', label='TCU067')
    # ida.plot_all()
    # # ida.plot_median()
    # multi_pushover.plot([
    #     'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-MMCUSER'
    # ])
    # plt.legend(loc='upper left')

    # ida.figure(xlim_max=0.02, intensity_measure='base_shear')
    # # ida.plot('TCU067', label='TCU067')
    # ida.plot_all()
    # # ida.plot_median()
    # multi_pushover.plot([
    #     'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
    # ])
    # plt.legend(loc='upper left')

    plt.show()


if __name__ == "__main__":
    main()
