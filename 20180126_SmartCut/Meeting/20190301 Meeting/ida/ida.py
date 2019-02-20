"""
ida data and function
"""
import os
import pickle

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


class IDA():
    """
    IDA data and function
    """

    def __init__(self, base_shear_path, story_drifts_path, story_displacements_path, earthquakes, stories):
        self.base_shear_path = base_shear_path
        self.story_drifts_path = story_drifts_path
        self.story_displacements_path = story_displacements_path
        self.earthquakes = earthquakes
        self.stories = stories
        self.base_shear = None
        self.story_drifts = None
        self.story_displacements = None
        self.intensity_measure = 'sa'
        self.damage_measure = 'story_drifts'

    def _story2level(self, df):
        for story in self.stories:
            df.loc[df['Story'] == story, 'StoryLevel'] = self.stories[story]

        return df

    def _init_damage(self, damage_measure='story_drifts'):
        """
        init time history story drifts or story displacements
        damage_measure='story_drifts' or story_displacements
        """
        if damage_measure == 'story_drifts':
            filepath = self.story_drifts_path
            sheet_name = 'Story Drifts'
        elif damage_measure == 'story_displacements':
            filepath = self.story_displacements_path
            sheet_name = 'Story Max Avg Displacements'

        pkl_file = f'{filepath} for IDA.pkl'

        if not os.path.exists(pkl_file):
            print("Reading excel...")

            read_file = f'{filepath}.xlsx'

            df = pd.read_excel(
                read_file, sheet_name=sheet_name, header=1, usecols=3, skiprows=[2])

            # convert story label to number
            df = self._story2level(df)

            # delete max and min string
            df.loc[:, 'Load Case/Combo'] = df['Load Case/Combo'].str[:-4]

            df = df.assign(
                StoryAndCase=lambda x: x['Story'] + ' ' + x['Load Case/Combo'])

            # combine max min
            df = df.groupby('StoryAndCase', as_index=False,
                            sort=False).agg('max')

            df['Load Case'], df['Scaled Factors'] = df['Load Case/Combo'].str.rsplit(
                '-', 1).str

            print("Creating pickle file ...")
            with open(pkl_file, 'wb') as f:
                pickle.dump(df, f, True)
            print("Done!")

        with open(pkl_file, 'rb') as f:
            df = pickle.load(f)

        if damage_measure == 'story_drifts':
            self.story_drifts = df
        elif damage_measure == 'story_displacements':
            self.story_displacements = df

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

            # delete max and min string
            df.loc[:, 'Load Case/Combo'] = df['Load Case/Combo'].str[:-4]

            # combine max min
            df = df.groupby('Load Case/Combo', as_index=False,
                            sort=False).agg('max')

            df['Load Case'], df['Scaled Factors'] = df['Load Case/Combo'].str.rsplit(
                '-', 1).str

            df.loc[:, 'Accel'] = np.abs(df['FX'] / df['FZ']) / 0.81

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

    def get_story_damage(self, damage_measure='story_drifts'):
        """
        get every step pushover story drifts or story displacements
        """
        if damage_measure == 'story_drifts':
            if self.story_drifts is None:
                self._init_damage(damage_measure)

            return self.story_drifts
        if damage_measure == 'story_displacements':
            if self.story_displacements is None:
                self._init_damage(damage_measure)

            return self.story_displacements

        return None

    def _get_damage_measure_column(self, damage_measure='story_drifts'):
        """
        return 'Drift' or 'Maximum'
        """
        if damage_measure == 'story_drifts':
            damage_measure_column = 'Drift'
        elif damage_measure == 'story_displacements':
            damage_measure_column = 'Maximum'
        return damage_measure_column

    def get_damage(self, damage_measure='story_drifts'):
        """
        condense story drift to max drift
        """
        column = self._get_damage_measure_column(damage_measure)

        story_damage = self.get_story_damage(damage_measure)

        damage = story_damage[story_damage.groupby(
            'Load Case/Combo')[column].transform(max) == story_damage[column]]

        damage = damage.drop_duplicates('Load Case/Combo')

        return damage

    def get_points(self, earthquake, damage_measure='story_drifts', intensity_measure='sa'):
        """
        get damage and intensity by loadcase and damage_measure
        """
        # 'Drift' or 'Maximum'
        column = self._get_damage_measure_column(damage_measure)
        damage = self.get_damage(damage_measure)
        intensity = self.earthquakes[earthquake][intensity_measure]

        damage = damage.loc[damage['Load Case'] == earthquake, :].copy()

        damage.loc[:, 'Scaled Factors'] = damage.loc[
            :, 'Scaled Factors'].astype('float64') * intensity

        damage = damage.sort_values(by=['Scaled Factors'])

        return damage[column].values, damage['Scaled Factors'].values

    def interp(self, damage_measure='story_drifts', intensity_measure='sa', num=1000):
        x = pd.DataFrame()
        y = pd.DataFrame()
        interpolation_x = pd.DataFrame()

        for earthquake in self.earthquakes:
            damage, intensity = self.get_points(
                earthquake, damage_measure, intensity_measure)

            # concat all drift and accel
            x = pd.concat(
                [x, pd.DataFrame({earthquake: damage})], axis=1)
            y = pd.concat(
                [y, pd.DataFrame({earthquake: intensity})], axis=1)

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

    def plot_all(self, *args, damage_measure='story_drifts', intensity_measure='sa', **kwargs):
        """
        plot ida in drift and acceleration by load case
        """
        for earthquake in self.earthquakes:
            damage, intensity = self.get_points(
                earthquake, damage_measure, intensity_measure)

            if not damage.size == 0:
                # plt.plot(xnew, f(xnew), label=earthquake, marker='.')
                plt.plot(damage, intensity, label=earthquake,
                         marker='.', *args, **kwargs)
            else:
                print(f'{earthquake} is not in data')

        plt.legend(loc='upper left')

    def plot(self, earthquake, *args,
             damage_measure='story_drifts', intensity_measure='sa', **kwargs):
        """
        plot pushover in drift and acceleration by load case
        """
        damage, intensity = self.get_points(
            earthquake, damage_measure, intensity_measure)
        plt.plot(damage, intensity, *args, **kwargs)

    def plot_median(self, *args,
                    damage_measure='story_drifts', intensity_measure='sa', **kwargs):
        """
        plot pushover in drift and acceleration by load case
        """
        damage, intensity = self.interp(
            damage_measure, intensity_measure, num=1000)
        plt.plot(damage.quantile(0.5, axis=1, interpolation='nearest'),
                 intensity, label='median IDA curve', *args, **kwargs)

    def figure(self,
               ylim_max=4, xlim_max=0.025,
               damage_measure='story_drifts', intensity_measure='sa',
               title='IDA versus Static Pushover for a 3-storey moment resisting frame'):
        """
        figure
        """
        plt.figure()
        plt.title(title)

        if damage_measure == 'story_drifts':
            plt.xlabel(r'Maximum interstorey drift ratio, $\theta_{max}$')
        elif damage_measure == 'story_displacements':
            plt.xlabel('Maximum displacement(mm)')

        if intensity_measure == 'sa':
            plt.ylabel(r'"first-mode"spectral acceleration $S_a(T_1$, 5%)(g)')
        elif intensity_measure == 'pga':
            plt.ylabel('Peak ground acceleration PGA(g)')

        if xlim_max is not None:
            plt.xlim(0, xlim_max)
        if ylim_max is not None:
            plt.ylim(0, ylim_max)


def _main():
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

    file_dir = os.path.dirname(os.path.abspath(__file__))

    ida = IDA(
        base_shear_path=file_dir+'/20190220 multi ida base shear',
        story_drifts_path=file_dir + '/20190214 multi ida story drifts',
        story_displacements_path=file_dir + '/20190214 multi ida displacement',
        earthquakes=earthquakes,
        stories=stories
    )

    print(ida.get_intensity())

    ida.figure(xlim_max=300, ylim_max=None,
               damage_measure='story_displacements')
    ida.plot_median(damage_measure='story_displacements')

    plt.show()

    # ida.plot('PUSHX-T', damage_measure='story_displacements',
    #   label = 'Static Pushover Curve')

    # plt.legend(loc='upper left')
    # plt.show()


if __name__ == "__main__":
    _main()
