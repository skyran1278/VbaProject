import time


class Clock():
    def __init__(self):
        self.start_time = None
        self.count = 1

    def _format_title(self, title):
        if title is None:
            return f'Lap {self.count}'
        return title

    def time(self, title=None):
        if self.start_time == None:
            self.start_time = time.time()
            print(f'--- {self._format_title(title)} ---')
        else:
            # print("--- %.3f seconds ---" % (time.time() - self.start_time))
            print(
                f'--- {time.time() - self.start_time} seconds ---\n')

            self.count += 1
            self.start_time = None


def main():
    clock = Clock()
    clock.time('3 points')
    clock.time('3 points')
    clock.time()
    clock.time()
    clock.time()


if __name__ == '__main__':
    main()
