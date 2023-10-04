'''
sokolov: An LLM-experimenting suite for my singular 'they' research. Will include data migration, cleanup, and checking features.

sokolov sokolov sokolov!
sokolov?
That's right.
'''

# standard libraries
import logging
import argparse


# project resources
from argparse_assets import handle_args



def main():
    logging.basicConfig(level=logging.NOTSET, format='INFO: %(message)s')
    args = handle_args()


if __name__ == "__main__":
    main()