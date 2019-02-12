import pandas as pd
import numpy as np

from single_IDA_points import single_IDA_points


def interp_IDAS(earthquakes, story_drifts, accel_unit='sa', num=1000):
    x = pd.DataFrame()
    y = pd.DataFrame()
    interpolation_x = pd.DataFrame()

    for earthquake in earthquakes:
        drifts, accelerations = single_IDA_points(
            earthquake, earthquakes, story_drifts, accel_unit)

        # concat all drift and accel
        x = pd.concat(
            [x, pd.DataFrame({earthquake: drifts})], axis=1)
        y = pd.concat(
            [y, pd.DataFrame({earthquake: accelerations})], axis=1)

    # scaled to same y
    interpolation_y = np.linspace(y.min().max(), y.max().min(), num=num)

    # interp nan, to delete nan
    x = x.interpolate()
    y = y.interpolate()

    # interpolate to same y to x
    for column in x:
        interpolation_x.loc[:, column] = np.interp(
            interpolation_y, y[column], x[column])

    return interpolation_x, interpolation_y
