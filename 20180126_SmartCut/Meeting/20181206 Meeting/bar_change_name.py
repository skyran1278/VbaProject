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

DATASET = pd.read_excel(SCRIPT_DIR + '/second_run_v1.xlsx',
                        sheet_name='beam_ld_added')

AB_7 = 3.871
AB_8 = 5.067
AB_10 = 8.143

TOP_SIZE = AB_10
BOT_SIZE = AB_8

X_NUM = 1200
X_NUM_3 = 400

START = 35
END = 525 - 37.5

SPAN = END - START
SPAN_3 = SPAN / 3


def conservative_cut(color):
    # Linear Conservative Cut
    # plot_bar(np.array([7, 3, 8]) * TOP_SIZE,
    #          np.array([11, 6, 10]) * BOT_SIZE, color=color)

    plot_bar_length(np.array([7, 3, 8]) * TOP_SIZE, [199.9883, 52.52333, 199.9883],
                    np.array([11, 6, 10]) * BOT_SIZE, [150.8333, 150.8333, 150.8333], color=color)


def no_etabs(color):
    # No ETABS
    plot_bar_length(np.array([9, 5, 9]) * TOP_SIZE, [229.125, 599.25, 229.125],
                    np.array([5, 4, 4]) * BOT_SIZE, [317.25, 564, 176.25], color=color)


def rcad(color):
    # RCAD
    plot_bar(np.array([7, 7, 8]) * TOP_SIZE, np.array([11, 10, 10])
             * BOT_SIZE, color=color, linewidth=4.0)


def linearcut(color):
    # # Linear Cut
    plot_bar_length(np.array([7, 6, 8]) * TOP_SIZE, [173.7, 119.9, 158.9],
                    np.array([11, 8, 10]) * BOT_SIZE, [143.7, 189.6, 119.2], color=color)

# =====================


def plot_bar(top_rebar, bot_rebar, color, linewidth=linewidth):
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


def real_sol(color):
    # Real Solution
    plt.plot(DATASET['StnLoc'] * 100,
             DATASET['BarTopNumLd'] * TOP_SIZE, color=color, linewidth=linewidth)
    plt.plot(DATASET['StnLoc'] * 100, -
             DATASET['BarBotNumLd'] * BOT_SIZE, color=color, linewidth=linewidth)


def conservative_sol(color):
    # Conservative Solution
    plt.plot(DATASET['StnLoc'] * 100,
             DATASET['BarTopNumSimpleLd'] * TOP_SIZE, color=color, linewidth=linewidth)
    plt.plot(DATASET['StnLoc'] * 100, -
             DATASET['BarBotNumSimpleLd'] * BOT_SIZE, color=color, linewidth=linewidth)


def etabs_demand(color):
    # ETABS Demand
    plt.plot(DATASET['StnLoc'] * 100, DATASET['AsTop']
             * 10000, color=color, linewidth=linewidth)
    plt.plot(DATASET['StnLoc'] * 100, -DATASET['AsBot']
             * 10000, color=color, linewidth=linewidth)


def etabs_to_addedld_sol():
    plt.figure()
    zero_line()

    etabs_demand(blue)

    real_sol(green)


def compare_RCAD():
    plt.figure()
    zero_line()

    etabs_demand(blue)
    real_sol(blue)

    # RCAD
    rcad(red)

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

    # # Linear Cut
    linearcut(green)


def conservative_flow():
    plt.figure()
    zero_line()

    etabs_demand(blue)

    conservative_sol(red)
    conservative_cut(green)


def linearcut_flow():
    plt.figure()
    zero_line()

    etabs_demand(blue)

    real_sol(red)
    linearcut(green)


def compare_real_to_conservative():
    plt.figure()
    zero_line()

    real_sol(red)
    conservative_cut(green)


conservative_flow()
linearcut_flow()
compare_real_to_conservative()
# etabs_to_addedld_sol()
compare_RCAD()
# no_etabs_enough_conservative()
# compare_linear_cut()


plt.show()
