'''
argparse_assets: Functions and types that I need for my argparse implementation.
'''

import os
import argparse

def dir_path(string) -> str:
    """
    Test if a given path exists on the machine.
    Used as ArgParser type.
    """
    if os.path.isdir(string):
        return string
    else:
        raise NotADirectoryError(string)


class MacroFile(Exception):
    pass


def file_path(string) -> str:
    """
    Test if a given file exists on the machine.
    Used as ArgParser type.
    """
    if os.path.isfile(string) and not string.endswith(".xlsm"):
        return string
    elif string.endswith(".xlsm"):
        raise MacroFile("Given filepath leads to an Excel macro file, which is not supported by Sokolov.")
    else:
        raise FileNotFoundError(string)   


def define_parser() -> argparse.ArgumentParser:
    """Define console argument parser."""
    parser = argparse.ArgumentParser(description="Keyword search comments from the Pushshift data dumps")

    # directories
    parser.add_argument('--inputfile', '-IF', type=file_path, required=True,
                        help="The file containing the data to be handled.")
    parser.add_argument('--output', '-O', type=dir_path, required=True,
                        help="The directory where search results will be saved to.")

    return parser


def handle_args() -> argparse.Namespace:
    """Handle argument-related edge cases by throwing meaningful errors."""
    parser = define_parser()
    args = parser.parse_args()
    
    return args