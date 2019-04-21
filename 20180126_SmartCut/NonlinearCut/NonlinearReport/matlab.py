"""
matlab data
"""
import os
import scipy.io as sio

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


class Matlab():
    """
    focus on matlab data
    """

    def __init__(self, mat_path):
        mat = sio.loadmat(mat_path)

        self.sa = (mat['scaled_factor'] * mat['sa']).flatten()

        self.triangle_displacement = (
            mat['sd_triangle'] * mat['PF1_phi']).flatten()
        self.uniform_displacement = (
            mat['sd_uniform'] * mat['PF1_phi']).flatten()
        self.power_displacement = (mat['sd_power'] * mat['PF1_phi']).flatten()

        self.triangle_sa = (mat['sa_triangle']).flatten()
        self.uniform_sa = (mat['sa_uniform']).flatten()
        self.power_sa = (mat['sa_power']).flatten()

    def plot(self, load_pattern, *args, **kwargs):
        """
        plot pushover in displacement and acceleration by load pattern
        """
        if load_pattern == 'triangle':
            plt.plot(self.triangle_displacement, self.sa, *args, **kwargs)
        elif load_pattern == 'uniform':
            plt.plot(self.uniform_displacement, self.sa, *args, **kwargs)
        elif load_pattern == 'power':
            plt.plot(self.power_displacement, self.sa, *args, **kwargs)

    def plot_all(self, *args, **kwargs):
        """
        plot all pushover in displacement and acceleration
        """
        plt.plot(self.triangle_displacement, self.triangle_sa,
                 label='Pushover Triangle', *args, **kwargs)
        plt.plot(self.uniform_displacement, self.uniform_sa,
                 label='Pushover Uniform', *args, **kwargs)
        plt.plot(self.power_displacement, self.power_sa,
                 label='Pushover Power', *args, **kwargs)
