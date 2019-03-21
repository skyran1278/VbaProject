class Materials:
    def __init__(self, data=None):
        if data is None:
            self.data = {}
        self.data = data

    def post(self, story, height):
        self.data[story] = height

    def get(self, story=None):
        if story is None:
            return self.data
        return self.data[story]
