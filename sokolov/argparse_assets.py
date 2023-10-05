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

# custom exceptions
class NotExcelFile(Exception):
    pass

class UnknownTask(Exception):
    pass


def task_type(string) -> str:
    known_tasks = ['preparation', 'experiment']
    if string not in known_tasks:
        raise UnknownTask(f'{string} is not a known task to Sokolov.')
    else:
        return string


def file_path(string) -> str:
    """
    Test if a given file exists on the machine.
    Used as ArgParser type.
    """
    if os.path.isfile(string):
        return string
    else:
        raise FileNotFoundError(string)
    
def excel_file(string) -> str:
    """
    Test if a given file exists on the machine.
    Used as ArgParser type.
    """
    if os.path.isfile(string) and string.endswith(".xlsx"):
        return string
    elif not string.endswith(".xlsx"):
        raise NotExcelFile(f"Given filepath, {string}, does not lead to a standard Excel file, which is the only file supported by Sokolov to date.")
    else:
        raise FileNotFoundError(string)


def define_parser() -> argparse.ArgumentParser:
    """Define console argument parser."""
    parser = argparse.ArgumentParser(description="Keyword search comments from the Pushshift data dumps")

    # directories
    parser.add_argument('--inputfile', '-IF', type=excel_file, required=True,
                        help="The file containing the data to be handled.")
    
    # tasks
    parser.add_argument('--task', '-T', type=task_type, required=True,
                        help="The task to be done with the file. Can be 'preparation' or 'experiment'.")

    # filters
    parser.add_argument('--length', '-L', type=int, required=False,
                        help= "Length filter for the comments. Any comment that's longer will be deleted from the file.")

    # experimenting
    parser.add_argument('--prompt', '-P', type=file_path, required=False,
                        help="The file containing the prompt design to be used.")
    parser.add_argument('--llm', '-LLM', type=str, required=False,
                        help="The LLM to be used for the experiment.")
    parser.add_argument('--moderation', '-M', action="store_true",
                        help="Send all comments to openAI moderation to check if they are acceptable to be used on GPT.")

    return parser


def handle_args() -> argparse.Namespace:
    """Handle argument-related edge cases by throwing meaningful errors."""
    parser = define_parser()
    args = parser.parse_args()

    with open(args.prompt, "r", encoding="utf-8") as infile:
        prompt_text = infile.read()
    
    args.prompt = prompt_text
    
    return args