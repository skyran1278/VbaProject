""" config from user. """
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = SCRIPT_DIR + '/data/'

config = {  # pylint: disable=invalid-name
    # 'design_path': DATA_DIR + '20190323 203316 SmartCut.xlsx',
    'design_path_test_v1': DATA_DIR + '20190327 173536 SmartCut.xlsx',
}
