"""
e2k model
"""
from utils.load_file import load_file


class E2k:
    def __init__(self, path):
        self.content = load_file(path)
