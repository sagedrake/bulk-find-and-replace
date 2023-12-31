import os
import glob
import powerpoint
import excel
import word
import eventlog

OLD = "Whale"  # word to find, beginning with a capital
NEW = "Shark"  # replacement word, beginning with a capital
FOLDER = "TestFiles"  # folder to find/replace within
LOG_OUTPUT_FILE = 'log.txt'  # file path for log output
FILE_SIZE_LIMIT = 1e9  # file size limit in bytes


def replace_contents_of_files():
    """
    Find all .docx, .pptx, and .xlsx files and edit their contents -- replacing the old word with the new word
    :return: None
    """
    files_to_edit = find_files()
    for filepath in files_to_edit['docx']:
        if not file_too_big(filepath):
            word.docx_find_and_replace(filepath, OLD, NEW)
    for filepath in files_to_edit['xlsx']:
        if not file_too_big(filepath):
            excel.xlsx_find_and_replace(filepath, OLD, NEW)
    for filepath in files_to_edit['pptx']:
        if not file_too_big(filepath):
            powerpoint.pptx_find_and_replace(filepath, OLD, NEW)


def file_too_big(filepath):
    """ Return True if the file at the given path exceeds the size limit, False otherwise. """
    size = os.path.getsize(filepath)
    if size > FILE_SIZE_LIMIT:
        eventlog.log_event("ERROR: Cannot edit contents of " + filepath + " because file exceeds 1Gb limit.")
        return True
    return False


def rename_files_and_folders(folder):
    """
    Recursively search through given directory and rename items containing the old word, replacing it with the new word.
    :param folder: The path of the directory to search within
    :return: None
    """
    for item in os.listdir(folder):
        # recursion to rename items in inner folders
        if os.path.isdir(folder + "/" + item):
            rename_files_and_folders(folder + "/" + item)

        if OLD.lower() in item.lower():
            new_name = str_find_and_replace(item, OLD, NEW)
            try:
                os.rename(folder + "/" + item, folder + "/" + new_name)
                eventlog.log_event("renamed " + item + " to " + new_name)
            except PermissionError:
                eventlog.log_event("ERROR: could not rename " + item + " because this file is already in use.")


def find_files():
    """ Return paths of all files ending in .docx, .pptx, or .xlsx in FOLDER or its subdirectories """
    found_files = {
        "docx": [],
        "pptx": [],
        "xlsx": []
    }

    for ext_to_find in found_files:
        found_files[ext_to_find] = glob.glob(FOLDER + '/**/*.' + ext_to_find, recursive=True)

    return found_files


def str_find_and_replace(string, old_word, new_word):
    """
    Replace instances of the given old word with the given new word in the given string
    :param string: The string to be searched within
    :param old_word: The word to search for
    :param new_word: The word to replace instances of the old word with
    :return: The inputted string with instances of the old word replaced with instances of the new word
    """
    # old -> new
    lowercase_replaced = string.replace(old_word.lower(), new_word.lower())
    # OLD -> NEW
    uppercase_replaced = lowercase_replaced.replace(old_word.upper(), new_word.upper())
    # Old -> New
    capitalized_replaced = uppercase_replaced.replace(old_word, new_word)

    return capitalized_replaced


def is_capitalized(text):
    """ Return True if text contains one uppercase letter followed by all lowercase letters, and False otherwise. """
    return text[:1].isupper() and text[1:].islower()


def __main__():
    """ If constants at top of file are valid, find and replace the specified words within the specified directory. """
    if not is_capitalized(OLD) or not is_capitalized(NEW):
        print("Both OLD and NEW must consist of one capital letter followed by all lowercase letters.")
        print("Your input does not meet this requirement. Please fix it and try again.")
        return

    if not os.path.isdir(FOLDER):
        print("The specified directory:" + FOLDER + " does not exist. Please fix this and try again.")
        return

    try:
        rename_files_and_folders(FOLDER)
        replace_contents_of_files()
    finally:
        eventlog.output_to_file(LOG_OUTPUT_FILE)


__main__()
