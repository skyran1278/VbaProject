from app.models.e2k import E2k


def test_materials(snapshot):
    path = 'D:/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190103 v3.0 3floor v16.e2k'

    e2k = E2k(path)
    print(e2k.materials)
    snapshot.assert_match(e2k.materials)
