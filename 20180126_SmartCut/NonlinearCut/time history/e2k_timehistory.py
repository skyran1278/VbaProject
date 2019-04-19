"""
generate function and loadcase e2k with peernga data
"""
import shlex


def put_timehistorys(time_historys, peernga_folder):
    """
    put NPTS, DT
    """
    for time_history in time_historys:
        with open(f'{peernga_folder}//{time_history}.AT2', encoding='big5') as f:
            words = shlex.split(f.readlines()[3])
            time_historys[time_history]['NPTS'] = int(words[1][:-1])
            time_historys[time_history]['DT'] = float(words[3])


def post_functions(time_historys, peernga_folder):
    """
    path folder same as time history

    @example
    FUNCTION "RSN169_IMPVALL.H_H-DLT262"  FUNCTYPE "HISTORY"
    FILE "path.AT2"  DATATYPE "EQUAL"  DT 0.01

    FUNCTION "RSN169_IMPVALL.H_H-DLT262"  HEADERLINES 4  POINTSPERLINE 5  FORMAT "FREE"
    """
    # post functions to list
    functions = []

    for time_history in time_historys:
        delta_t = time_historys[time_history]['DT']

        functions.append(
            f'FUNCTION "{time_history}"  FUNCTYPE "HISTORY"  '
            f'FILE "{peernga_folder}\\{time_history}.AT2"  '
            f'DATATYPE "EQUAL"  DT {delta_t}\n'
        )

        functions.append(
            f'FUNCTION "{time_history}"  HEADERLINES 4  POINTSPERLINE 5  FORMAT "FREE"\n'
        )

    return functions


def post_loadcases(time_historys, period, initial_condition, direction):
    """
    loadcase

    @example
    LOADCASE "timehistory"  TYPE  "Nonlinear Direct Integration History"
    INITCOND  "PUSHDLLL"  MODALCASE  "Modal"  MASSSOURCE  "Previous"

    LOADCASE "timehistory"  ACCEL  "U1"  FUNC  "chichi_TCU052_max"  SF  1

    LOADCASE "timehistory"  NUMBEROUTPUTSTEPS  100 OUTPUTSTEPSIZE  0.1

    LOADCASE "timehistory"  PRODAMPTYPE  "Period"  T1  0.344 DAMP1  0.05 T2  0.088 DAMP2  0.05

    LOADCASE "timehistory"  MODALDAMPTYPE  "Constant"  CONSTDAMP  0.05
    CONSIDERMAXMODALFREQ  "Yes"  MAXCONSIDEREDMODALFREQ  100

    LOADCASE "timehistory"  USEEVENTSTEPPING  "No"
    """
    loadcases = []

    for time_history in time_historys:
        factors = time_historys[time_history]['FACTORS']
        number_output_steps = time_historys[time_history]['NPTS']
        delta_t = time_historys[time_history]['DT']

        for factor in factors:
            name = f'{time_history}-{factor}'

            # G to m/s2
            factor = factor * 9.81

            loadcases.append(
                f'LOADCASE "{name}"  TYPE  "Nonlinear Direct Integration History"  '
                f'INITCOND  "{initial_condition}"  MODALCASE  "Modal"  MASSSOURCE  "Previous"\n'
            )

            loadcases.append(
                f'LOADCASE "{name}"  ACCEL  "{direction}"  '
                f'FUNC  "{time_history}"  SF  {factor}\n'
            )
            loadcases.append(
                f'LOADCASE "{name}"  NUMBEROUTPUTSTEPS  {number_output_steps} '
                f'OUTPUTSTEPSIZE  {delta_t}\n'
            )
            loadcases.append(
                f'LOADCASE "{name}"  PRODAMPTYPE  "Period"  T1  {period[0]} DAMP1  0.05 '
                f'T2  {period[1]} DAMP2  0.05\n'
            )
            loadcases.append(
                f'LOADCASE "{name}"  MODALDAMPTYPE  "Constant"  CONSTDAMP  0.05 '
                f'CONSIDERMAXMODALFREQ  "Yes"  MAXCONSIDEREDMODALFREQ  100 \n'
            )
            loadcases.append(f'LOADCASE "{name}"  USEEVENTSTEPPING  "No"\n')

    return loadcases


def main():
    """
    test
    """
    import os

    # global
    script_folder = os.path.dirname(os.path.abspath(__file__))

    peernga_folder = script_folder + '\\PEERNGARecords_Unscaled'

    direction = 'U1'

    initial_condition = 'PUSHDLLL'

    # different by model
    period = [0.039, 0.039 / 10]

    time_historys = {
        'RSN169_IMPVALL.H_H-DLT262': {
            'FACTORS': [1, 2, 3]
        },
        'RSN953_NORTHR_MUL009':  {
            'FACTORS': [1, 2]
        },
    }

    put_timehistorys(time_historys, peernga_folder)
    functions = post_functions(time_historys, peernga_folder)
    loadcases = post_loadcases(
        time_historys, period, initial_condition, direction
    )

    with open(script_folder + '/e2k_timehistory.e2k', mode='w', encoding='big5') as f:
        f.writelines(functions)
        f.write('\n\n\n')
        f.writelines(loadcases)


if __name__ == "__main__":
    main()
