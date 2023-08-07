import openpyxl as xl


def xlsx_find_and_replace(filepath, old_word, new_word):
    """
    Replace instances of the given old word with the given new word in sheet titles, cell contents of given Excel file
    :param filepath: The name of the file to be edited
    :param old_word: The word to be searched for, assumed to be capitalized (e.g. 'Shark' not 'shark')
    :param new_word: The word to replace old_word with, assumed to be capitalized (e.g. 'Whale' not 'WHALE')
    :return: None
    """
    wb = xl.load_workbook(filepath)

    for sheet in wb.worksheets:
        sheet.title = str_find_and_replace(sheet.title, old_word, new_word)
        for row in range(1, sheet.max_row + 1):
            for col in range(1, sheet.max_column + 1):
                cell = sheet.cell(row, col)
                if isinstance(cell.value, str):
                    cell.value = str_find_and_replace(cell.value, old_word, new_word)

    wb.save(filepath)


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


xlsx_find_and_replace("TestFiles/SHARK.xlsx", "Whale", "Shark")
