"""
pushover data and function
"""
import os
import pickle

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


class Pushover():
    """
    pushover data and function
    """

    def __init__(self, base_shear_path, story_drifts_path, stories, loadcases):
        self.base_shear_path = base_shear_path
        self.story_drifts_path = story_drifts_path
        self.stories = stories
        self.loadcases = loadcases
        self.story_drifts = None
        self.base_shear = None

    def _story2level(self, df):
        for story in self.stories:
            df.loc[df['Story'] == story, 'StoryLevel'] = self.stories[story]

        return df

    def get_story_drifts(self):
        """
        get every step pushover story drifts
        """
        if not self.story_drifts:
            self._init_story_drifts()

        return self.story_drifts

    def _init_story_drifts(self):
        """
        get every step pushover story drifts
        """
        pkl_file = f'{self.story_drifts_path} for pushover.pkl'

        if not os.path.exists(pkl_file):
            print("Reading excel...")

            read_file = f'{self.story_drifts_path}.xlsx'

            df = pd.read_excel(
                read_file, sheet_name='Story Drifts', header=1, usecols=3, skiprows=[2])

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

        self.story_drifts = df

    def get_max_drifts(self):
        """
        condense story drift to max drift
        """
        if self.story_drifts is None:
            self._init_story_drifts()

        max_drifts = self.story_drifts[self.story_drifts.groupby(
            'Load Case/Combo')['Drift'].transform(max) == self.story_drifts['Drift']]

        max_drifts = max_drifts.drop_duplicates('Load Case/Combo')

        # max_drifts = max_drifts.reset_index(drop=True)

        # print(max_drifts.head(30))

        # max_drifts = self.story_drifts.groupby(
        #     'Load Case/Combo', as_index=False, sort=False)['Drift'].max()

        return max_drifts

    def _init_base_shear(self):
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

            df.loc[:, 'Accel'] = np.abs(df['FX'] / df['FZ'])

            print("Creating pickle file ...")
            with open(pkl_file, 'wb') as f:
                pickle.dump(df, f, True)
            print("Done!")

        with open(pkl_file, 'rb') as f:
            df = pickle.load(f)

        self.base_shear = df

    def get_base_shear(self):
        """
        get pushover base shear and acceleration
        """
        if self.base_shear is None:
            self._init_base_shear()
        return self.base_shear

    def get_loadcase_drift_and_accel(self, loadcase):
        """
        get drift and accel by load case
        """
        drifts = self.get_max_drifts()
        base_shear = self.get_base_shear()

        drifts = drifts.loc[drifts['Load Case'] == loadcase, :].copy()
        drifts.loc[:, 'Step'] = drifts['Step'].astype('float64')
        drifts = drifts.sort_values(by=['Step'])

        base_shear = base_shear.loc[base_shear['Load Case']
                                    == loadcase, :].copy()
        base_shear.loc[:, 'Step'] = base_shear.loc[:, 'Step'].astype('float64')
        base_shear = base_shear.sort_values(by=['Step'])

        return drifts['Drift'].values, base_shear['Accel'].values

    def plot_in_drift_and_accel(self, loadcase):
        """
        plot pushover in drift and acceleration by load case
        """
        drifts, base_shear = self.get_loadcase_drift_and_accel(loadcase)
        plt.plot(drifts, base_shear)


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

    pushover = Pushover(story_drifts_path=file_dir + '/20190212 pushover story drifts',
                        base_shear_path=file_dir + '/20190212 pushover base shear',
                        stories=stories, loadcases=loadcases)

    pushover.plot_in_drift_and_accel('PUSHX-T')
    plt.show()


if __name__ == "__main__":
    _main()
