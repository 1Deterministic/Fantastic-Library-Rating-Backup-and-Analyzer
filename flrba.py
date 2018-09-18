import _tag

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
        print("Reading files...")
        # walks the received folder searching for mp3 files
        for (dirpath, dirnames, filenames) in os.walk(target_path):
            for f in filenames:
                if os.path.splitext(f)[1] == ".mp3":
                    tag = _tag.Tag(os.path.join(dirpath, f))
                    # if the file was loaded correctly, add it to the database
                    if not tag.title == "<none>":
                        try:
                            cursor.executescript("insert or replace into Track (title, album, artist, genre, lenght, rating) values (" +
                                                 "\"" + tag.title + "\", " +
                                                 "\"" + tag.album + "\", " +
                                                 "\"" + tag.artist + "\", " +
                                                 "\"" + tag.genre + "\", " +
                                                 "" + str(tag.lenght) + ", " +
                                                 "" + str(tag.rating) + ");")
                        except sqlite3.Error as er:
                            print ('er:', er.message)
                        #except:
                        #    print("Error: could not save " + tag.to_string())
        # at the end, commits the changes to the file
        try:
            connection.commit()
            print("Done!")
        except:
            print("Error: could not save info to the database file.")
            sys.exit()

    # if restore_ratings was received instead
    elif restore_ratings_to:
        files = []
        print("Reading files...")

        # walks the received folder searching for mp3 files
        for (dirpath, dirnames, filenames) in os.walk(target_path):
            for f in filenames:
                if os.path.splitext(f)[1] == ".mp3":
                    tag = _tag.Tag(os.path.join(dirpath, f))
                    # if the file was loaded correctly, add it to the list
                    if not tag.title == "<none>":
                        files.append(tag)
        print("Done!")

        # gets information about the tracks saved in the database file
        print("Reading saved files...")
        cursor.execute("select * from Track;")
        print("Done!")

        print("Restoring ratings...")
        for line in cursor.fetchall():
            # 0 = title, 1 = album, 2 = artist, 3 = genre, 4 = lenght, 5 = rating
            # for every file identified in the folder received
            for f in files:
                # will restore the rating if title, album and title match or if they are similar enough (when approximation was received)
                if (approximation and (fuzz.ratio(line[0], f.title) > 90 and fuzz.ratio(line[1], f.album) > 90 and fuzz.ratio(line[2], f.artist) > 90)) or (line[0] == f.title and line[1] == f.album and line[2] == f.artist):
                    f.rating = line[5];
                    if not f.write():
                        print("Error: could not restore rating to " + f.to_string())
                    break
        print("Done!")


    # if export_excel was received
    if export_excel:
        # creates an apropriate data structure to store the information
        data = collections.defaultdict(lambda: collections.defaultdict(lambda: collections.defaultdict(dict)))

        # gets the info stored in the database file
        print("Reading saved files...")
        cursor.execute("select * from Track;")
        print("Done!")

        # fills the dict with the information received from the query
        print("Building data structure...")
        for line in cursor.fetchall():
            # 0 = title; 1 = album; 2 = artist; 3 = genre; 4 = lenght; 5 = rating
            data[line[2]][line[1]][line[0]]["rating"] = line[5]
            data[line[2]][line[1]][line[0]]["genre"] = line[3]
            data[line[2]][line[1]][line[0]]["lenght"] = line[4]
        print("Done!")

        # creates the xlsx file
        workbook = xlsxwriter.Workbook("music_database.xlsx")
        # adds a new sheet to the file
        worksheet = workbook.add_worksheet("all tracks info")

        # writes the xlsx file
        print("Writing to .xlsx file...")

        # all information ==============================================================================================
        # starting row of the spreadsheet (headers index)
        row = 1
        # writes the headers of the table
        worksheet.write("A" + str(row), "Artist")
        worksheet.write("B" + str(row), "Album")
        worksheet.write("C" + str(row), "Title")
        worksheet.write("D" + str(row), "Rating")
        worksheet.write("E" + str(row), "Lenght")
        worksheet.write("F" + str(row), "Genre")
        row += 1

        # inserts the values in the spreadsheet, starting from the row 2
        for artist in data.keys():
            for album in data[artist]:
                for title in data[artist][album]:
                    worksheet.write("A" + str(row), artist)
                    worksheet.write("B" + str(row), album)
                    worksheet.write("C" + str(row), title)
                    worksheet.write("D" + str(row), data[artist][album][title]["rating"])
                    worksheet.write("E" + str(row), data[artist][album][title]["lenght"])
                    worksheet.write("F" + str(row), data[artist][album][title]["genre"])
                    row += 1


        # average track rating per artist ==============================================================================
        worksheet = workbook.add_worksheet("average track rating per artist")
        row = 1
        # writes the headers of the table
        worksheet.write("A" + str(row), "Artist")
        worksheet.write("B" + str(row), "Average Rating")
        row += 1

        # inserts the values in the spreadsheet, starting from the row 2
        for artist in data.keys():
            worksheet.write("A" + str(row), artist)

            sum = 0
            tracks = 0
            for album in data[artist]:
                for title in data[artist][album]:
                    # consider only valid ratings (0 is not rated yet)
                    if int(data[artist][album][title]["rating"]) > 0:
                        sum += int(data[artist][album][title]["rating"])
                        tracks += 1

            worksheet.write("B" + str(row), sum / tracks if tracks > 0 else 0)
            row += 1


        # average album rating per artist ==============================================================================
        worksheet = workbook.add_worksheet("average album rating per artist")
        row = 1
        # writes the headers of the table
        worksheet.write("A" + str(row), "Artist")
        worksheet.write("B" + str(row), "Album")
        worksheet.write("C" + str(row), "Average Rating")
        row += 1

        # inserts the values in the spreadsheet, starting from the row 2
        for artist in data.keys():
            worksheet.write("A" + str(row), artist)
            for album in data[artist]:
                worksheet.write("B" + str(row), album)

                sum = 0
                tracks = 0
                for title in data[artist][album]:
                    if int(data[artist][album][title]["rating"]) > 0:
                        sum += int(data[artist][album][title]["rating"])
                        tracks += 1

                worksheet.write("C" + str(row), sum / tracks if tracks > 0 else 0)
                row += 1


        # number of 5-star tracks per artist ===========================================================================
        worksheet = workbook.add_worksheet("# of 5-star tracks per artist")
        row = 1
        # writes the headers of the table
        worksheet.write("A" + str(row), "Artist")
        worksheet.write("B" + str(row), "Number of 5-star tracks")
        row += 1

        # inserts the values in the spreadsheet, starting from the row 2
        for artist in data.keys():
            worksheet.write("A" + str(row), artist)

            sum = 0
            for album in data[artist]:
                for title in data[artist][album]:
                    if int(data[artist][album][title]["rating"]) == 5:
                        sum += 1

            worksheet.write("B" + str(row), sum)
            row += 1


        """# lenght range and avg rating ===================================================================================
        worksheet = workbook.add_worksheet("lenght range and avg rating")
        row = 1
        # writes the headers of the table
        worksheet.write("A" + str(row), "Range")
        worksheet.write("B" + str(row), "Number of tracks")
        worksheet.write("C" + str(row), "Average rating")
        row += 1

        # inserts the values in the spreadsheet, starting from the row 2
        ranges = ["less than 1 minute", "1-2 minutes", "2-3 minutes", "3-4 minutes", "4-5 minutes", "more than 5 minutes"]
        ranges_sums = [0, 0, 0, 0, 0, 0]
        ranges_rating_total = [0, 0, 0, 0, 0, 0]

        for artist in data.keys():
            for album in data[artist]:
                for title in data[artist][album]:
                    if int(data[artist][album][title]["lenght"]) < 60: ranges_sums[0] += 1; ranges_rating_total[0] += int(data[artist][album][title]["rating"])
                    elif int(data[artist][album][title]["lenght"]) < 120: ranges_sums[1] += 1; ranges_rating_total[1] += int(data[artist][album][title]["rating"])
                    elif int(data[artist][album][title]["lenght"]) < 180: ranges_sums[2] += 1; ranges_rating_total[2] += int(data[artist][album][title]["rating"])
                    elif int(data[artist][album][title]["lenght"]) < 240: ranges_sums[3] += 1; ranges_rating_total[3] += int(data[artist][album][title]["rating"])
                    elif int(data[artist][album][title]["lenght"]) < 300: ranges_sums[4] += 1; ranges_rating_total[4] += int(data[artist][album][title]["rating"])
                    elif int(data[artist][album][title]["lenght"]) >= 300: ranges_sums[5] += 1; ranges_rating_total[5] += int(data[artist][album][title]["rating"])

        index = 0
        for r in ranges:
            worksheet.write("A" + str(row), r)
            worksheet.write("B" + str(row), ranges_sums[index])
            worksheet.write("C" + str(row), ranges_rating_total[index] / ranges_sums[index] if ranges_sums[index] > 0 else 0)
            index += 1
            row += 1
        """

        # closes the file
        workbook.close()
        print("Done!")