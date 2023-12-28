from datetime import datetime

""" This module provides logging functionality for events that only have a date/time of occurrence and a description """


class Event:
    """ Represents an event with a date/time of occurrence and a description. """

    def __init__(self, description):
        """
        Create an event with the given description and the current date & time.
        :param description: A description of the event
        """
        self.description = description
        self.date_logged = datetime.now()

    def __str__(self):
        """ Return a string representation of the event including the date, time, and description. """
        return str(self.date_logged) + ": " + self.description


events = []


def log_event(description):
    """
    Add an event with the given description to the log.
    :param description: A description of the event
    :return: None
    """
    events.append(Event(description))
    print(description)


def output_to_file(filepath):
    """
    Output all logged events to the file with the given path, with each event on a new line.
    :param filepath: The path of the file to be outputted to
    :return: None
    """
    file = open(filepath, "w")
    for e in events[:]:
        file.write(str(e) + "\n")
    file.close()
