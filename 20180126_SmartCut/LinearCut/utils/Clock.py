import time

from colorama import init
from colorama import Fore
from colorama import Back
from colorama import Style

init()


class Clock():
    def __init__(self):
        self.start_time = None
        self.title = None
        self.count = 1

    def _format_title(self, title):
        if title is None:
            return f'{Fore.BLACK}{Back.WHITE}Lap {self.count}{Style.RESET_ALL}'

        return f'\n{Fore.BLACK}{Back.WHITE}{title}{Style.RESET_ALL}'

    def time(self, title=None):
        if title is not None:
            self.title = title

        if self.start_time is None:
            self.start_time = time.time()
            print(f'{self._format_title(self.title)}')

        else:
            # print("%.3f seconds" % (time.time() - self.start_time))
            print(
                f'{Fore.BLUE}{round(time.time() - self.start_time, 4)}{Style.RESET_ALL} seconds')

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
