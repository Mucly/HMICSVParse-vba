import os
import sys


if __name__ == "__main__":
    res = sys.argv[1]
    target = sys.argv[2]

    cwd = os.getcwd()
    for file in os.listdir(cwd):
        head, tail = os.path.splitext(file)
        if res in head:
            head = head.replace(res, target)
            os.rename(file, head + "." + tail)
    os.system('pause')
