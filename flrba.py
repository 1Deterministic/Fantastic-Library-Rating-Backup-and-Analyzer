import _save_from
import _restore_ratings_to
import _export_excel
import _tag
import _counter

import collections
import sqlite3
import os
import sys
import logging

import xlsxwriter
from fuzzywuzzy import fuzz

if __name__ == "__main__":
    # hides the logs from eyed3
    logging.getLogger("eyed3").setLevel(logging.CRITICAL)

    # execution parameters
    save_from = False
    restore_ratings_to = False
    approximation = False
    export_excel = False
    target_path = ""

    # checks for save_from received
    try:
        if (sys.argv.index("save_from")):
            save_from = True
            try:
                if os.path.isdir(sys.argv[sys.argv.index("save_from") + 1]):
                    target_path = sys.argv[sys.argv.index("save_from") + 1]
                else:
                    print("Path received is not a folder")
                    sys.exit()
            except IndexError:
                print("Missing folder path")
                sys.exit()
    except ValueError:
        save_from = False


    # checks for restore_ratings_to received
    try:
        if (sys.argv.index("restore_ratings_to")):
            restore_ratings_to = True
            try:
                if os.path.isdir(sys.argv[sys.argv.index("restore_ratings_to") + 1]):
                    target_path = sys.argv[sys.argv.index("restore_ratings_to") + 1]
                else:
                    print("Path received is not a folder")
                    sys.exit()
            except IndexError:
                print("Missing folder path")
                sys.exit()
            try:
                if (sys.argv.index("approximation")):
                    approximation = True
            except ValueError:
                approximation = False
    except ValueError:
        restore_ratings_to = False


    # checks for report received
    try:
        if (sys.argv.index("export_excel")):
            export_excel = True
    except ValueError:
        export_excel = False


    # prevents saving and restoring at the same time
    if save_from and restore_ratings_to:
        print("Cannot save and restore at the same time")
        sys.exit()


    # loads the database if it already exists or create a new one if it doesn't
    try:
        connection = sqlite3.connect("database.db")
        cursor = connection.cursor()
        cursor.executescript(open("database_sql.sql", "r").read())
        connection.commit()
    except:
        print("Error: could not create the database file.")
        sys.exit()


    # if save_from was received
    if save_from:
        _save_from.save(target_path, cursor)

    # if restore_ratings was received instead
    elif restore_ratings_to:
        _restore_ratings_to.restore(target_path, cursor, approximation)

    # if export_excel was received
    if export_excel:
        _export_excel.export(cursor)

    # commits any changes to the database file
    try:
        print("Saving changes to the database file...")
        connection.commit()
        print("Done!")
    except Exception:
        print("Error: could not save to the database file")
        sys.exit()