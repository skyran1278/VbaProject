"""
load file
"""


def load_file(path):
    """
    load file by path\n
        return file data
    """
    with open(path, encoding='big5') as filename:
        data = filename.readlines()
        data = [x.strip() for x in data]

    return data
