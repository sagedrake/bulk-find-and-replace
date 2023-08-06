import os

OLD = "Whale"  # word to find, beginning with a capital
NEW = "Shark"  # replacement word, beginning with a capital
FOLDER = 'C:/Users/sagew/OneDrive/Documents/CS Projects/Find and Replace/WhalesAreCool'  # folder to find/replace within


# Rename all files and folders in the given directory
def rename_items(folder):
    for item in os.listdir(folder):
        # recursion to rename items in inner folders
        if os.path.isdir(folder + "/" + item):
            rename_items(folder + "/" + item)

        # old -> new
        if OLD.lower() in item:
            os.rename(folder + "/" + item, folder + "/" + item.replace(OLD.lower(), NEW.lower()))
            print("renamed " + item + " to " + item.replace(OLD.lower(), NEW.lower()))
        # Old -> New
        elif OLD in item:
            os.rename(folder + "/" + item, folder + "/" + item.replace(OLD, NEW))
            print("renamed " + item + " to " + item.replace(OLD, NEW))
        # OLD -> NEW
        elif OLD.upper() in item:
            os.rename(folder + "/" + item, folder + "/" + item.replace(OLD.upper(), NEW.upper()))
            print("renamed " + item + " to " + item.replace(OLD.upper(), NEW.upper()))


rename_items(FOLDER)

