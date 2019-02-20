"""
Plotlib
"""
import matplotlib.pyplot as plt


class Plotlib():
    """
    Plotlib
    intensity_measure
    damage_measure
    """

    def __init__(self):
        self.intensity_measure = None
        self.damage_measure = None

    def figure(self, ylim_max=None, xlim_max=None,
               damage_measure='story_drifts', intensity_measure='sa',
               title='IDA versus Static Pushover for a 3-storey moment resisting frame'):
        """
        figure
        """
        plt.figure()
        plt.title(title)

        self.intensity_measure = intensity_measure
        self.damage_measure = damage_measure

        if damage_measure == 'story_drifts':
            plt.xlabel(r'Maximum interstorey drift ratio, $\theta_{max}$')
        elif damage_measure == 'story_displacements':
            plt.xlabel('Maximum displacement(mm)')

        if intensity_measure == 'sa':
            plt.ylabel(r'"first-mode"spectral acceleration $S_a(T_1$, 5%)(g)')
        elif intensity_measure == 'pga':
            plt.ylabel('Peak ground acceleration PGA(g)')
        elif intensity_measure == 'base_shear':
            plt.ylabel('Base shear(tonf)')

        if xlim_max is not None:
            plt.xlim(0, xlim_max)
        if ylim_max is not None:
            plt.ylim(0, ylim_max)

    def show(self):
        pass
