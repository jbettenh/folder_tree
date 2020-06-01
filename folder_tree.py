#! python 3
# This will list out all folders and subfolders for a given directory.

import argparse
import os


def dir_path(string):
    if os.path.isdir(string):
        return string
    else:
        raise NotADirectoryError(string)


def arguments():
    parser = argparse.ArgumentParser(description='Parser')
    parser.add_argument('--path', type=dir_path, help='This is the path to the folder')

    return parser.parse_args()


def main():
    args = arguments()

    for foldername in os.walk(args.path):
        print(f'Folder: {foldername}')


if __name__ == '__main__':
    main()