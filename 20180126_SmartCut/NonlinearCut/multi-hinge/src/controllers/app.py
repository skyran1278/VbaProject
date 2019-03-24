"""
entry
"""
from src.models.e2k import E2k
from src.models.design import Design
from src.controllers.get_points import get_points


def main():
    """
    test
    """
    # pylint: disable=line-too-long

    design = Design(
        'D:/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190323 203316 SmartCut.xlsx')

    e2k = E2k(
        'D:/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190103 v3.0 3floor v16.e2k')

    for index in range(0, design.get_len(), 4):
        points = get_points(index, design, e2k)


if __name__ == "__main__":
    main()
