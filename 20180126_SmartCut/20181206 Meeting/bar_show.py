import os
import sys
import pickle
import random

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))


INPUT_FILE = 'first_run_v3_kaohsiung'
INDEX = random.randrange(0, 24, 4)
INDEX = 12
print(INDEX)

green = np.array([26, 188, 156]) / 256
blue = np.array([52, 152, 219]) / 256
red = np.array([233, 88, 73]) / 256
orange = np.array([230, 126, 34]) / 256
gray = np.array([0.5, 0.5, 0.5])
background = np.array([247, 247, 247]) / 256

linewidth = 2.0

if not os.path.exists(f'{SCRIPT_DIR}/{INPUT_FILE}.pkl'):
    DATASET = pd.read_excel(SCRIPT_DIR + f'/{INPUT_FILE}.xlsx',
                            sheet_name='beam_ld_added')
    BEAM_3 = pd.read_excel(SCRIPT_DIR + f'/{INPUT_FILE}.xlsx',
                           sheet_name='三點斷筋', header=[0, 1], usecols=20)
    BEAM_CON = pd.read_excel(SCRIPT_DIR + f'/{INPUT_FILE}.xlsx',
                             sheet_name='傳統斷筋', header=[0, 1], usecols=20)

    print("Creating pickle file ...")
    with open(f'{SCRIPT_DIR}/{INPUT_FILE}.pkl', 'wb') as f:
        pickle.dump((DATASET, BEAM_3, BEAM_CON), f, True)
    print("Done!")

with open(f'{SCRIPT_DIR}/{INPUT_FILE}.pkl', 'rb') as f:
    DATASET, BEAM_3, BEAM_CON = pickle.load(f)

DATASET = DATASET.loc[(DATASET['Story'] == BEAM_3.at[INDEX, ('樓層', 'Unnamed: 0_level_1')]) & (
    DATASET['BayID'] == BEAM_3.at[INDEX, ('編號', 'Unnamed: 1_level_1')])]

REBARS = {
    "#2": 0.32258,
    "#3": 0.709676,
    "#4": 1.29032,
    "#5": 1.999996,
    "#6": 2.838704,
    "#7": 3.87096,
    "#8": 5.096764,
    "#9": 6.4516,
    "#10": 8.193532,
    "#11": 10.0645
}

# AB_7 = 3.871
# AB_8 = 5.067
# AB_10 = 8.143
TOP_INDEX = INDEX
TOP_INDEX_2 = INDEX + 1
BOT_INDEX = INDEX + 3
BOT_INDEX_2 = BOT_INDEX - 1

TOP_SIZE = REBARS[BEAM_3.at[TOP_INDEX, ('主筋', '左')].split('-')[1]]
BOT_SIZE = REBARS[BEAM_3.at[BOT_INDEX, ('主筋', '左')].split('-')[1]]

X_NUM = 1200
X_NUM_3 = 400

START = BEAM_3.at[INDEX, ('支承寬', '左')]
END = BEAM_3.at[INDEX, ('梁長', 'Unnamed: 15_level_1')] - \
    BEAM_3.at[INDEX, ('支承寬', '右')]

SPAN = END - START
SPAN_3 = SPAN / 3


def sum_rebar(beam, bar_1, bar_2, loc):
    if beam.at[bar_2, ('主筋', loc)] == 0:
        return int(beam.at[bar_1, ('主筋', loc)].split('-')[0])
    return int(beam.at[bar_1, ('主筋', loc)].split('-')[0]) + int(beam.at[bar_2, ('主筋', loc)].split('-')[0])


def conservative_cut(color):
    top_rebar = np.array([
        sum_rebar(BEAM_CON, TOP_INDEX, TOP_INDEX_2, '左'),
        sum_rebar(BEAM_CON, TOP_INDEX, TOP_INDEX_2, '中'),
        sum_rebar(BEAM_CON, TOP_INDEX, TOP_INDEX_2, '右')
    ]) * TOP_SIZE

    top_length = np.array([BEAM_CON.at[TOP_INDEX, ('長度', '左')],
                           BEAM_CON.at[TOP_INDEX, ('長度', '中')],
                           BEAM_CON.at[TOP_INDEX, ('長度', '右')]])

    bot_rebar = np.array([
        sum_rebar(BEAM_CON, BOT_INDEX, BOT_INDEX_2, '左'),
        sum_rebar(BEAM_CON, BOT_INDEX, BOT_INDEX_2, '中'),
        sum_rebar(BEAM_CON, BOT_INDEX, BOT_INDEX_2, '右')
    ]) * BOT_SIZE

    bot_length = np.array([BEAM_CON.at[BOT_INDEX, ('長度', '左')],
                           BEAM_CON.at[BOT_INDEX, ('長度', '中')],
                           BEAM_CON.at[BOT_INDEX, ('長度', '右')]])

    plot_bar_length(top_rebar, top_length, bot_rebar, bot_length, color=color)

    return np.sum(top_rebar * top_length), np.sum(bot_rebar * bot_length)


def no_etabs(color):
    # No ETABS
    plot_bar_length(np.array([9, 5, 9]) * TOP_SIZE, [229.125, 599.25, 229.125],
                    np.array([5, 4, 4]) * BOT_SIZE, [317.25, 564, 176.25], color=color)


def rcad(color):
    # RCAD
    plot_bar(np.array([7, 7, 8]) * TOP_SIZE, np.array([11, 10, 10])
             * BOT_SIZE, color=color, linewidth=4.0)


def linearcut(color):
    # Linear Cut

    top_rebar = np.array([
        sum_rebar(BEAM_3, TOP_INDEX, TOP_INDEX_2, '左'),
        sum_rebar(BEAM_3, TOP_INDEX, TOP_INDEX_2, '中'),
        sum_rebar(BEAM_3, TOP_INDEX, TOP_INDEX_2, '右')
    ]) * TOP_SIZE

    top_length = np.array([BEAM_3.at[TOP_INDEX, ('長度', '左')],
                           BEAM_3.at[TOP_INDEX, ('長度', '中')],
                           BEAM_3.at[TOP_INDEX, ('長度', '右')]])

    bot_rebar = np.array([
        sum_rebar(BEAM_3, BOT_INDEX, BOT_INDEX_2, '左'),
        sum_rebar(BEAM_3, BOT_INDEX, BOT_INDEX_2, '中'),
        sum_rebar(BEAM_3, BOT_INDEX, BOT_INDEX_2, '右')
    ]) * BOT_SIZE

    bot_length = np.array([BEAM_3.at[BOT_INDEX, ('長度', '左')],
                           BEAM_3.at[BOT_INDEX, ('長度', '中')],
                           BEAM_3.at[BOT_INDEX, ('長度', '右')]])

    plot_bar_length(top_rebar, top_length, bot_rebar, bot_length, color=color)

    return np.sum(top_rebar * top_length), np.sum(bot_rebar * bot_length)

    # # Linear Cut
    # plot_bar_length(
    #     np.array([
    #         sum_rebar(BEAM_3, TOP_INDEX, TOP_INDEX_2, '左'),
    #         sum_rebar(BEAM_3, TOP_INDEX, TOP_INDEX_2, '中'),
    #         sum_rebar(BEAM_3, TOP_INDEX, TOP_INDEX_2, '右')
    #     ]) * TOP_SIZE,
    #     [BEAM_3.at[TOP_INDEX, ('長度', '左')],
    #      BEAM_3.at[TOP_INDEX, ('長度', '中')],
    #      BEAM_3.at[TOP_INDEX, ('長度', '右')]],
    #     np.array([
    #         sum_rebar(BEAM_3, BOT_INDEX, BOT_INDEX_2, '左'),
    #         sum_rebar(BEAM_3, BOT_INDEX, BOT_INDEX_2, '中'),
    #         sum_rebar(BEAM_3, BOT_INDEX, BOT_INDEX_2, '右')
    #     ]) * BOT_SIZE,
    #     [BEAM_3.at[BOT_INDEX, ('長度', '左')],
    #      BEAM_3.at[BOT_INDEX, ('長度', '中')],
    #      BEAM_3.at[BOT_INDEX, ('長度', '右')]],
    #     color=color)

# =====================


def plot_bar(top_rebar, bot_rebar, color, linewidth=linewidth):
    x = np.empty((X_NUM // 3, 1))

    plt.plot(np.linspace(START, END, X_NUM), np.concatenate((np.full_like(
        x, top_rebar[0]), np.full_like(x, top_rebar[1]), np.full_like(x, top_rebar[2]))), color=color, linewidth=linewidth)
    plt.plot(np.linspace(START, END, X_NUM), np.concatenate((np.full_like(
        x, -bot_rebar[0]), np.full_like(x, -bot_rebar[1]), np.full_like(x, -bot_rebar[2]))), color=color, linewidth=linewidth)


def plot_bar_length(top_rebar, top_length, bot_rebar, bot_length, color):
    plt.plot([START, (START + top_length[0]), (START + top_length[0]), (START + top_length[0] +
                                                                        top_length[1]), END - top_length[2], END], np.repeat(top_rebar, 2), color=color, linewidth=linewidth)
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


def verticalline():
    plt.axvline((END - START) / 4 + START, linestyle='--', color=gray)
    plt.axvline((END - START) / 3 + START, linestyle='--', color=gray)
    plt.axvline((END - START) / 3 * 2 + START, linestyle='--', color=gray)
    plt.axvline((END - START) / 4 * 3 + START, linestyle='--', color=gray)


def horizontalline():
    plt.axhline(2 * TOP_SIZE, linestyle='--', color=gray)
    plt.axhline(-2 * BOT_SIZE, linestyle='--', color=gray)


def conservative_flow():
    plt.figure()
    zero_line()

    verticalline()
    horizontalline()

    etabs_demand(blue)

    conservative_sol(red)
    conservative_cut(green)


def linearcut_flow():
    plt.figure()
    zero_line()

    horizontalline()

    etabs_demand(blue)

    real_sol(red)
    linearcut(green)


def compare_real_to_conservative():
    plt.figure()
    zero_line()

    horizontalline()

    real_sol(blue)
    top_con, bot_con = conservative_cut(red)
    top_3, bot_3 = linearcut(green)

    print(f'con top: {top_con} bot: {bot_con}')
    print(f'3 top: {top_3} bot: {bot_3}')
    print(f'top: {top_3 / top_con} bot: {bot_3 / bot_con}')


conservative_flow()
linearcut_flow()
compare_real_to_conservative()
# etabs_to_addedld_sol()
# compare_RCAD()
# no_etabs_enough_conservative()
# compare_linear_cut()


plt.show()
