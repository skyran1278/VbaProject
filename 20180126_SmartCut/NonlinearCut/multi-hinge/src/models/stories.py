class Stories:
    def __init__(self, stories=None):
        if stories is None:
            self.data = {}
        self.data = stories

    def post(self, story, height):
        self.data[story] = height

    def get(self, story=None):
        if story is None:
            return self.data
        return self.data[story]
