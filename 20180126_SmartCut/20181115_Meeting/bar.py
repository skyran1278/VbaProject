import os
import sys

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

green = np.array([26, 188, 156]) / 256
blue = np.array([52, 152, 219]) / 256
red = np.array([233, 88, 73]) / 256
orange = np.array([230, 126, 34]) / 256
gray = np.array([0.5, 0.5, 0.5])
background = np.array([247, 247, 247]) / 256

linewidth = 2.0

DATASET = pd.read_excel(SCRIPT_DIR + '/first_run_v3.xlsx',
                        sheet_name='beam_ld_added')

X_NUM = 1200
X_NUM_3 = 400

START = 45
END = 1160 - 57.5

SPAN = 1057.5
SPAN_3 = SPAN / 3

AB_7 = 3.871
AB_8 = 5.067


def plot_bar(top_rebar, bot_rebar, color):
    x = np.empty((X_NUM // 3, 1))

    plt.plot(np.linspace(START, END, X_NUM), np.concatenate((np.full_like(
        x, top_rebar[0]), np.full_like(x, top_rebar[1]), np.full_like(x, top_rebar[2]))), color=color, linewidth=linewidth)
    plt.plot(np.linspace(START, END, X_NUM), np.concatenate((np.full_like(
        x, -bot_rebar[0]), np.full_like(x, -bot_rebar[1]), np.full_like(x, -bot_rebar[2]))), color=color, linewidth=linewidth)


def plot_bar_length(top_rebar, top_length, bot_rebar, bot_length, color):
    plt.plot([START, START + top_length[0], START + top_length[0], START + top_length[0] +
              top_length[1], END - top_length[2], END], np.repeat(top_rebar, 2), color=color, linewidth=linewidth)
    plt.plot([START, START + bot_length[0], START + bot_length[0], START + bot_length[0] +
              bot_length[1], END - bot_length[2], END], -np.repeat(bot_rebar, 2), color=color, linewidth=linewidth)


def zero_line():
    # 基準線
    plt.plot([START, END], [0, 0], color=gray, linewidth=linewidth)


def conservative_cut(color):
    # Linear Conservative Cut
    plot_bar(np.array([9, 3, 9]) * AB_7,
             np.array([5, 4, 4]) * AB_7, color=color)


def real_sol(color):
    # Real Solution
    plt.plot(DATASET['StnLoc'] * 100,
             DATASET['BarTopNumLd'] * AB_7, color=color, linewidth=linewidth)
    plt.plot(DATASET['StnLoc'] * 100, -
             DATASET['BarBotNumLd'] * AB_7, color=color, linewidth=linewidth)


def no_etabs(color):
    # No ETABS
    plot_bar_length(np.array([9, 5, 9]) * AB_7, [229.125, 599.25, 229.125],
                    np.array([5, 4, 4]) * AB_7, [317.25, 564, 176.25], color=color)


def etabs_to_addedld_sol():
    plt.figure()
    zero_line()

    # ETABS Demand
    plt.plot(DATASET['StnLoc'] * 100, DATASET['AsTop']
             * 10000, color=blue, linewidth=linewidth)
    plt.plot(DATASET['StnLoc'] * 100, -DATASET['AsBot']
             * 10000, color=blue, linewidth=linewidth)

    real_sol(green)


def compare_RCAD():
    plt.figure()
    zero_line()

    real_sol(blue)

    # RCAD
    plot_bar(np.array([7, 3, 7]) * AB_8, np.array([4, 3, 4])
             * AB_8, color=red)

    conservative_cut(green)


def no_etabs_enough_conservative():
    plt.figure()
    zero_line()

    real_sol(blue)

    conservative_cut(red)

    no_etabs(green)


def compare_linear_cut():
    plt.figure()
    zero_line()

    real_sol(blue)

    conservative_cut(red)

    # Linear Cut
    # plot_bar_length(np.array([9, 4, 9]) * AB_7, [320, 521.4, 216.1],
    #                 np.array([5, 4, 4]) * AB_7, [160, 720.7, 176.8], green)
    # # Linear Cut
    plot_bar_length(np.array([9, 4, 9]) * AB_7, [250, 532.5, 275],
                    np.array([5, 4, 4]) * AB_7, [170, 484, 403.5], color=green)


etabs_to_addedld_sol()
compare_RCAD()
no_etabs_enough_conservative()
compare_linear_cut()

plt.show()
