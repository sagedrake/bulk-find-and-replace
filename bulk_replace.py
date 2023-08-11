import os
import glob
import powerpoint
import excel
import word

OLD = "Whale"  # word to find, beginning with a capital
NEW = "Shark"  # replacement word, beginning with a capital
FOLDER = 'TestFiles'  # folder to find/replace within


def bulk_find_and_replace():
    rename_files_and_folders(FOLDER)

    files_to_edit = find_files()
    for filepath in files_to_edit['docx']: 
        word.docx_find_and_replace(filepath, OLD, NEW)
        print("completed find and replace within " + filepath)
    for filepath in files_to_edit['xlsx']:
        excel.xlsx_find_and_replace(filepath, OLD, NEW)
        print("completed find and replace within " + filepath)
    for filepath in files_to_edit['pptx']:
        powerpoint.pptx_find_and_replace(filepath, OLD, NEW)
        print("completed find and replace within " + filepath)


def rename_files_and_folders(folder):
    for item in os.listdir(folder):
        # recursion to rename items in inner folders
        if os.path.isdir(folder + "/" + item):
            rename_files_and_folders(folder + "/" + item)

        # old -> new
        if OLD.lower() in item.lower():
            new_name = str_find_and_replace(item, OLD, NEW)
            os.rename(folder + "/" + item, folder + "/" + new_name)
            print("renamed " + item + " to " + new_name)


def find_files():
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

bulk_find_and_replace()

