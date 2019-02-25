"""
pushover data and function
"""
import os
import pickle

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from plotlib import EnhancePlotlib


class Pushover():
    """
    pushover data and function
    """

    def __init__(self, path, stories):
        super().__init__()
        self.base_shear_path = path['base_shear_path']
        self.story_drifts_path = path['story_drifts_path']
        self.story_displacements_path = path['story_displacements_path']
        self.stories = stories

        self.story_drifts = None
        self.story_displacements = None
        self.base_shear = None

        self.plotlib = EnhancePlotlib()

    def __getattr__(self, attr):
        return getattr(self.plotlib, attr)

    def _story2level(self, df):
        for story in self.stories:
            df.loc[df['Story'] == story, 'StoryLevel'] = self.stories[story]

        return df

    def get_story_damage(self):
        """
        get every step pushover story drifts or story displacements
        """
        if self.plotlib.damage_measure == 'story_drifts':
            if self.story_drifts is None:
                self._init_damage()
            return self.story_drifts

        if self.plotlib.damage_measure == 'story_displacements':
            if self.story_displacements is None:
                self._init_damage()
            return self.story_displacements

        return None

    def _init_damage(self):
        """
        init every step pushover story drifts or story displacements
        damage_measure='story_drifts' or story_displacements
        """
        if self.plotlib.damage_measure == 'story_drifts':
            filepath = self.story_drifts_path
            sheet_name = 'Story Drifts'
        elif self.plotlib.damage_measure == 'story_displacements':
            filepath = self.story_displacements_path
            sheet_name = 'Story Max Avg Displacements'

        pkl_file = f'{filepath} for pushover.pkl'

        if not os.path.exists(pkl_file):
            print("Reading excel...")

            read_file = f'{filepath}.xlsx'

            df = pd.read_excel(
                read_file, sheet_name=sheet_name, header=1, usecols=3, skiprows=[2])

            # convert story label to number
            df = self._story2level(df)

            # StoryAndCase = Story + Load Case/Combo
            df = df.assign(
                StoryAndCase=lambda x: x['Story'] + ' ' + x['Load Case/Combo'])

            # split Load Case/Combo to load case and step
            df.loc[:, 'Load Case'], df.loc[:, 'Step'] = df['Load Case/Combo'].str.rsplit(
                ' ', 1).str

            print("Creating pickle file ...")
            with open(pkl_file, 'wb') as f:
                pickle.dump(df, f, True)
            print("Done!")

        with open(pkl_file, 'rb') as f:
            df = pickle.load(f)

        if self.plotlib.damage_measure == 'story_drifts':
            self.story_drifts = df
        elif self.plotlib.damage_measure == 'story_displacements':
            self.story_displacements = df

    def _get_dm_col(self):
        """
        return 'Drift' or 'Maximum'
        """
        if self.plotlib.damage_measure == 'story_drifts':
            return 'Drift'
        if self.plotlib.damage_measure == 'story_displacements':
            return 'Maximum'

    def get_damage(self):
        """
        condense story drift to max drift
        """
        column = self._get_dm_col()

        story_damage = self.get_story_damage()

        damage = story_damage[story_damage.groupby(
            'Load Case/Combo')[column].transform(max) == story_damage[column]]

        damage = damage.drop_duplicates('Load Case/Combo')

        return damage

    def _init_intensity(self):
        """
        get pushover base shear and acceleration
        """
        pkl_file = f'{self.base_shear_path} for pushover.pkl'

        if not os.path.exists(pkl_file):
            print("Reading excel...")

            read_file = f'{self.base_shear_path}.xlsx'

            df = pd.read_excel(
                read_file, sheet_name='Base Reactions', header=1, usecols=3, skiprows=[2])

            # split Load Case/Combo to load case and step
            df.loc[:, 'Load Case'], df.loc[:, 'Step'] = df['Load Case/Combo'].str.rsplit(
                ' ', 1).str

            df.loc[:, 'FX'] = np.abs(df['FX'])
            df.loc[:, 'Accel'] = df['FX'] / 299.5941 / 0.81

            print("Creating pickle file ...")
            with open(pkl_file, 'wb') as f:
                pickle.dump(df, f, True)
            print("Done!")

        with open(pkl_file, 'rb') as f:
            df = pickle.load(f)

        self.base_shear = df

    def get_intensity(self):
        """
        get pushover base shear and acceleration
        """
        if self.base_shear is None:
            self._init_intensity()
        return self.base_shear

    def _get_im_col(self):
        """
        return 'Drift' or 'Maximum'
        """
        if self.plotlib.intensity_measure == 'sa':
            return 'Accel'
        if self.plotlib.intensity_measure == 'base_shear':
            return 'FX'

    def get_points(self, loadcase):
        """
        get damage and intensity by loadcase and damage_measure
        """
        damage = self.get_damage()
        intensity = self.get_intensity()

        damage = damage.loc[damage['Load Case'] == loadcase, :].copy()
        damage.loc[:, 'Step'] = damage['Step'].astype('float64')
        damage = damage.sort_values(by=['Step'])

        intensity = intensity.loc[
            intensity['Load Case'] == loadcase, :].copy()
        intensity.loc[:, 'Step'] = intensity.loc[:, 'Step'].astype('float64')
        intensity = intensity.sort_values(by=['Step'])

        return (
            damage[self._get_dm_col()].values,
            intensity[self._get_im_col()].values
        )

    def plot(self, loadcases, *args, **kwargs):
        """
        plot pushover in drift and acceleration by load case
        """
        if isinstance(loadcases, str):
            loadcase = loadcases
            damage, intensity = self.get_points(loadcase)
            plt.plot(damage, intensity, marker='.', *args, **kwargs)

        else:
            for loadcase in loadcases:
                damage, intensity = self.get_points(loadcase)
                plt.plot(damage, intensity, marker='.',
                         label=loadcase, *args, **kwargs)

        plt.legend(loc='upper left')


def _main():
    stories = {
        'RF': 4,
        '3F': 3,
        '2F': 2,
    }

    loadcases = [
        'PUSHX-T', 'PUSHX-U', 'PUSHX-P', 'PUSHX-1', 'PUSHX-2', 'PUSHX-3', 'PUSHX-MMC',
        'PUSHX-1USER', 'PUSHX-2USER', 'PUSHX-3USER', 'PUSHX-MMCUSER'
    ]

    file_dir = os.path.dirname(os.path.abspath(__file__))

    path = {
        'base_shear_path': file_dir + '/20190214 multi pushover base shear',
        'story_drifts_path': file_dir + '/20190214 multi pushover story drifts',
        'story_displacements_path': file_dir + '/20190214 multi pushover displacement'
    }

    pushover = Pushover(path=path, stories=stories)

    pushover.figure(xlim_max=0.025, intensity_measure='sa')
    pushover.plot('PUSHX-T', label='Static Pushover Curve')
    plt.legend(loc='upper left')

    pushover.figure(xlim_max=0.025, intensity_measure='base_shear')
    pushover.plot(loadcases)
    plt.legend(loc='upper left')

    plt.show()


if __name__ == "__main__":
    _main()
