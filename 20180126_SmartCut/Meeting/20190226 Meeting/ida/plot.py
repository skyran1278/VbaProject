"""
plot ida
"""
import os

# import pandas as pd
# import numpy as np
import matplotlib.pyplot as plt
# from scipy.interpolate import interp1d

from pushover import Pushover
# from dataset import dataset_ida_storydrifts, dataset_ida_storydisp
# from interp_IDAS import interp_IDAS
# from plot_single_IDA import plot_single_IDA
# from plot_multi_IDAS import plot_multi_IDAS
# from plot_fractiles import plot_fractiles, plot_fractiles_log
# from plot_capacity_rule import plot_DM_rule, plot_IM_rule
# from plot_normal_versus_multi import plot_normal_versus_multi, plot_normal_versus_multi_log
# from plot_ida_pushover import plot_median_idas_and_pushover, plot_multi_idas_and_pushover
from ida import IDA

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

    ida.figure(xlim_max=0.025, intensity_measure='pga')
    ida.plot_all()
    ida.plot_median()
    plt.legend(loc='upper left')

    ida.figure(xlim_max=0.025, intensity_measure='sa')
    # ida.plot_all()
    ida.plot_median()
    multi_pushover.plot([
        'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-MMCUSER'
    ])
    plt.legend(loc='upper left')

    ida.figure(xlim_max=0.025, intensity_measure='sa')
    # ida.plot_all()
    ida.plot_median()
    multi_pushover.plot([
        'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
    ])
    plt.legend(loc='upper left')

    ida.figure(xlim_max=0.025, intensity_measure='base_shear')
    # ida.plot('TCU067', label='TCU067')
    ida.plot_all()
    # ida.plot_median()
    multi_pushover.plot([
        'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-MMCUSER'
    ])
    plt.legend(loc='upper left')

    ida.figure(xlim_max=0.025, intensity_measure='base_shear')
    # ida.plot('TCU067', label='TCU067')
    ida.plot_all()
    # ida.plot_median()
    multi_pushover.plot([
        'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
    ])
    plt.legend(loc='upper left')

    plt.show()

    # ida.figure(xlim_max=0.02, ylim_max=3.5)
    # ida.plot_all()

    # multi_pushover.plot([
    #     'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-1', 'PUSHX-2', 'PUSHX-3',
    #     'PUSHX-MMC', 'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
    # ], marker='.')

    # ida.figure(xlim_max=0.02, ylim_max=2.5)
    # ida.plot_median()

    # multi_pushover.plot([
    #     'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-MMCUSER'
    # ], marker='.')

    # ida.figure(xlim_max=0.02, ylim_max=2.5)
    # ida.plot_median()

    # multi_pushover.plot([
    #     'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
    # ], marker='.')

    # ida.figure(xlim_max=150, ylim_max=3.5,
    #            damage_measure='story_displacements')
    # ida.plot_all(damage_measure='story_displacements')

    # multi_pushover.plot([
    #     'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-1', 'PUSHX-2', 'PUSHX-3',
    #     'PUSHX-MMC', 'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
    # ], marker='.', damage_measure='story_displacements')

    # ida.figure(xlim_max=150, ylim_max=2.5,
    #            damage_measure='story_displacements')
    # ida.plot_median(damage_measure='story_displacements')

    # multi_pushover.plot([
    #     'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-MMCUSER'
    # ], marker='.', damage_measure='story_displacements')

    # ida.figure(xlim_max=150, ylim_max=2.5,
    #            damage_measure='story_displacements')
    # ida.plot_median(damage_measure='story_displacements')

    # multi_pushover.plot([
    #     'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
    # ], marker='.', damage_measure='story_displacements')

    # plt.show()


if __name__ == "__main__":
    main()


# multi_story_drifts = dataset_ida_storydrifts(
#     '20190214 multi ida story drifts', stories)
# # print(multi_story_drifts.head())
# # plot_single_IDA('TCU067', earthquakes, multi_story_drifts, ylim_max=3)

# # plot_multi_IDAS(earthquakes, multi_story_drifts,
# #                 ylim_max=None, xlim_max=None, accel_unit='pga')
# # plot_multi_IDAS(earthquakes, multi_story_drifts,
# #                 ylim_max=None, xlim_max=0.025, accel_unit='sa')

# # pushover.plot_all_loadcases()

# plot_multi_idas_and_pushover(
#     earthquakes, multi_story_drifts, multi_pushover, loadcases=[
#         'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-1', 'PUSHX-2',
#         'PUSHX-3', 'PUSHX-MMC', 'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
#     ], ylim_max=4, xlim_max=0.02)
# plot_median_idas_and_pushover(
#     earthquakes, multi_story_drifts, multi_pushover, loadcases=[
#         'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-MMCUSER'
#     ], ylim_max=2.5, xlim_max=0.02)
# plot_median_idas_and_pushover(
#     earthquakes, multi_story_drifts, multi_pushover, loadcases=[
#         'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
#     ], ylim_max=2.5, xlim_max=0.02)


# multi_story_disp = dataset_ida_storydisp(
#     '20190214 multi ida story displacement', stories)
# # print(multi_story_disp.head())
# # plot_single_IDA('TCU067', earthquakes, multi_story_disp, ylim_max=3)

# # plot_multi_IDAS(earthquakes, multi_story_disp,
# #                 ylim_max=None, xlim_max=None, accel_unit='pga')
# # plot_multi_IDAS(earthquakes, multi_story_disp,
# #                 ylim_max=None, xlim_max=0.025, accel_unit='sa')

# # pushover.plot_all_loadcases()

# plot_multi_idas_and_pushover(
#     earthquakes, multi_story_disp, multi_pushover, loadcases=[
#         'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-1', 'PUSHX-2',
#         'PUSHX-3', 'PUSHX-MMC', 'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
#     ], ylim_max=4, xlim_max=0.02)
# plot_median_idas_and_pushover(
#     earthquakes, multi_story_disp, multi_pushover, loadcases=[
#         'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-MMCUSER'
#     ], ylim_max=2.5, xlim_max=0.02)
# plot_median_idas_and_pushover(
#     earthquakes, multi_story_disp, multi_pushover, loadcases=[
#         'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
#     ], ylim_max=2.5, xlim_max=0.02)
# plot_median_idas_and_single_pushover(
#     earthquakes, multi_story_drifts, multi_pushover, 'PUSHX-T', ylim_max=None, xlim_max=0.025)
# plot_fractiles(earthquakes, multi_story_drifts, ylim_max=2, accel_unit='pga')
# plot_fractiles_log(earthquakes, multi_story_drifts,
#                    ylim_max=2, accel_unit='pga')

# plot_fractiles(earthquakes, multi_story_drifts,
#                ylim_max=4, xlim_max=0.025, accel_unit='sa')
# plot_fractiles_log(earthquakes, multi_story_drifts,
#                    ylim_max=4, xlim_max=0.25, accel_unit='sa')


# def median_ida_points(earthquakes, multi_story_drifts, accel_unit='sa'):
#     multi_x, multi_y = interp_IDAS(
#         earthquakes, multi_story_drifts, accel_unit='sa')
#     multi_x = multi_x.quantile(
#         0.5, axis=1, interpolation='nearest').values

#     return multi_x, multi_y

# plot_DM_rule(multi_x, multi_y, ylim_max=4,
#              xlim_max=0.25, accel_unit='sa', C_DM=0.02)
# plot_IM_rule(multi_x, multi_y, ylim_max=4,
#              xlim_max=0.25, accel_unit='sa', C_IM=2.03)


# normal_story_drifts = dataset_ida_storydrifts(
#     '20190214 normal ida story drifts', stories)

# normal_pushover = Pushover(
#     story_drifts_path=os.path.join(
#         file_dir, '20190214 normal pushover story drifts'),
#     base_shear_path=os.path.join(
#         file_dir, '20190214 normal pushover base shear'),
#     stories=stories
# )

# plot_multi_idas_and_pushover(
#     earthquakes, normal_story_drifts, normal_pushover, loadcases=[
#         'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-1', 'PUSHX-2',
#         'PUSHX-3', 'PUSHX-MMC', 'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
#     ], ylim_max=4, xlim_max=0.02)
# plot_median_idas_and_pushover(
#     earthquakes, normal_story_drifts, normal_pushover, loadcases=[
#         'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-MMCUSER'
#     ], ylim_max=2.5, xlim_max=0.02)
# plot_median_idas_and_pushover(
#     earthquakes, normal_story_drifts, normal_pushover, loadcases=[
#         'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
#     ], ylim_max=2.5, xlim_max=0.02)

# normal_scwb_story_drifts = dataset_ida_storydrifts(
#     '20190214 normal scwb ida story drifts', stories)

# normal_scwb_pushover = Pushover(
#     story_drifts_path=os.path.join(
#         file_dir, '20190214 normal scwb pushover story drifts'),
#     base_shear_path=os.path.join(
#         file_dir, '20190214 normal scwb pushover base shear'),
#     stories=stories
# )

# plot_multi_idas_and_pushover(
#     earthquakes, normal_scwb_story_drifts, normal_scwb_pushover, loadcases=[
#         'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-1', 'PUSHX-2',
#         'PUSHX-3', 'PUSHX-MMC', 'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
#     ], ylim_max=4, xlim_max=0.02)
# plot_median_idas_and_pushover(
#     earthquakes, normal_scwb_story_drifts, normal_scwb_pushover, loadcases=[
#         'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-MMCUSER'
#     ], ylim_max=2.5, xlim_max=0.02)
# plot_median_idas_and_pushover(
#     earthquakes, normal_scwb_story_drifts, normal_scwb_pushover, loadcases=[
#         'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
#     ], ylim_max=2.5, xlim_max=0.02)

# multi_scwb_story_drifts = dataset_ida_storydrifts(
#     '20190214 multi scwb ida story drifts', stories)

# multi_scwb_pushover = Pushover(
#     story_drifts_path=os.path.join(
#         file_dir, '20190214 multi scwb pushover story drifts'),
#     base_shear_path=os.path.join(
#         file_dir, '20190214 multi scwb pushover base shear'),
#     stories=stories
# )

# plot_multi_idas_and_pushover(
#     earthquakes, multi_scwb_story_drifts, multi_scwb_pushover, loadcases=[
#         'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-1', 'PUSHX-2',
#         'PUSHX-3', 'PUSHX-MMC', 'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
#     ], ylim_max=4, xlim_max=0.02)
# plot_median_idas_and_pushover(
#     earthquakes, multi_scwb_story_drifts, multi_scwb_pushover, loadcases=[
#         'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-MMCUSER'
#     ], ylim_max=2.5, xlim_max=0.02)
# plot_median_idas_and_pushover(
#     earthquakes, multi_scwb_story_drifts, multi_scwb_pushover, loadcases=[
#         'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
#     ], ylim_max=2.5, xlim_max=0.02)
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

# for loadcase in loadcases:
#     pushover.plot_in_drift_and_accel(loadcase)
# plot_normal_versus_multi(earthquakes, normal_story_drifts,
#                          multi_story_drifts, ylim_max=3, xlim_max=0.025, accel_unit='sa')
# plot_normal_versus_multi_log(earthquakes, normal_story_drifts,
#                              multi_story_drifts, ylim_max=3, xlim_max=0.025, accel_unit='sa')
# plot_normal_versus_multi(earthquakes, normal_story_drifts,
#                          multi_story_drifts, ylim_max=1.25, xlim_max=0.025, accel_unit='pga')
# plot_normal_versus_multi_log(earthquakes, normal_story_drifts,
#                              multi_story_drifts, ylim_max=1.25, xlim_max=0.025, accel_unit='pga')
