import time


class Clock():
    def __init__(self):
        self.start_time = None

    def time(self):
        if self.start_time == None:
            self.start_time = time.time()
        else:
            print("--- %s seconds ---" % (time.time() - self.start_time))
            self.start_time = None


def main():
    clock = Clock()
    clock.time()
    clock.time()
    clock.time()
    clock.time()
    clock.time()


if __name__ == '__main__':
    main()
