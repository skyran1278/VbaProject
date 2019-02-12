def single_IDA_points(earthquake, earthquakes, story_drifts, accel_unit='sa'):
    accel = earthquakes[earthquake][accel_unit]

    earthquake_drift = story_drifts.loc[story_drifts['Load Case']
                                        == earthquake, :]

    max_drift = earthquake_drift.groupby(
        'Scaled Factors', as_index=False, sort=False)['Drift'].max()

    max_drift.loc[:, 'Scaled Factors'] = max_drift.loc[:,
                                                       'Scaled Factors'].astype('float64') * accel

    max_drift = max_drift.sort_values(by=['Scaled Factors'])

    return max_drift['Drift'].values, max_drift['Scaled Factors'].values
