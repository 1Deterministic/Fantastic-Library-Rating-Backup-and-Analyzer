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
        workbook = xlsxwriter.Workbook("analytics.xlsx")
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
        # writes the headers of the page
        worksheet.write("A" + str(row), "Artist")
        worksheet.write("B" + str(row), "Average Rating")
        row += 1

        # gets the sum of valid ratings and the total of tracks
        rating_sum_counter = _counter.Counter()
        tracks_number_counter = _counter.Counter()
        for artist in data.keys():
            for album in data[artist]:
                for title in data[artist][album]:
                    # consider only valid ratings (0 means it's not rated yet)
                    if int(data[artist][album][title]["rating"]) > 0:
                        rating_sum_counter.push(artist, int(data[artist][album][title]["rating"]))
                        tracks_number_counter.push(artist)

        # gets the averages per artist
        average_ratings = dict()
        for n in range(0, len(rating_sum_counter.entries)):
            average_ratings[rating_sum_counter.entries[n].name] = rating_sum_counter.entries[n].number / tracks_number_counter.entries[n].number if tracks_number_counter.entries[n].number > 0 else 0

        # gets the sorted list
        sorted_average_ratings = sorted(average_ratings.items(), key=lambda value: value[1], reverse=True)

        # fills the spreadsheet
        data_initial_row = row
        for a in sorted_average_ratings:
            worksheet.write("A" + str(row), a[0])
            worksheet.write("B" + str(row), a[1])
            row += 1
        data_final_row = row

        # creates the bar chart
        chart = workbook.add_chart({'type': 'bar'})
        chart.add_series({
            "name": "average track rating",
            'categories': "='average track rating per artist'!A" + str(data_initial_row) + ":A" + str(data_final_row),
            'values':"='average track rating per artist'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

        chart.set_y_axis({'name': 'Artist'})
        chart.set_x_axis({'name': 'Average Rating'})
        worksheet.insert_chart('C1', chart, {'x_scale': 2, 'y_scale': 5})


        # average track rating per album ===============================================================================
        worksheet = workbook.add_worksheet("average track rating per album")
        row = 1
        # writes the headers of the page
        worksheet.write("A" + str(row), "Album")
        worksheet.write("B" + str(row), "Average Rating")
        row += 1

        # gets the sum of valid ratings and the total of tracks
        rating_sum_counter = _counter.Counter()
        tracks_number_counter = _counter.Counter()
        for artist in data.keys():
            for album in data[artist]:
                for title in data[artist][album]:
                    # consider only valid ratings (0 means it's not rated yet)
                    if int(data[artist][album][title]["rating"]) > 0:
                        rating_sum_counter.push(album, int(data[artist][album][title]["rating"]))
                        tracks_number_counter.push(album)

        # gets the averages per artist
        average_ratings = dict()
        for n in range(0, len(rating_sum_counter.entries)):
            average_ratings[rating_sum_counter.entries[n].name] = rating_sum_counter.entries[n].number / tracks_number_counter.entries[n].number if tracks_number_counter.entries[n].number > 0 else 0

        # gets the sorted list
        sorted_average_ratings = sorted(average_ratings.items(), key=lambda value: value[1], reverse=True)

        # fills the spreadsheet
        data_initial_row = row
        for a in sorted_average_ratings:
            worksheet.write("A" + str(row), a[0])
            worksheet.write("B" + str(row), a[1])
            row += 1
        data_final_row = row

        # creates the bar chart
        chart = workbook.add_chart({'type': 'bar'})
        chart.add_series({
            "name": "average track rating",
            'categories': "='average track rating per album'!A" + str(data_initial_row) + ":A" + str(data_final_row),
            'values':"='average track rating per album'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

        chart.set_y_axis({'name': 'Album'})
        chart.set_x_axis({'name': 'Average Rating'})
        worksheet.insert_chart('C1', chart, {'x_scale': 2, 'y_scale': 2})


        # number of 5-star tracks per artist ===========================================================================
        worksheet = workbook.add_worksheet("# of 5-star tracks per artist")
        row = 1
        # writes the headers of the table
        worksheet.write("A" + str(row), "Artist")
        worksheet.write("B" + str(row), "Number of 5-star tracks")
        row += 1

        counter = _counter.Counter()
        for artist in data.keys():
            for album in data[artist]:
                for title in data[artist][album]:
                    if int(data[artist][album][title]["rating"]) == 5:
                        counter.push(artist)

        sorted_number = sorted(counter.entries, key=lambda value: value.number, reverse=True)

        data_initial_row = row
        for c in sorted_number:
            worksheet.write("A" + str(row), c.name)
            worksheet.write("B" + str(row), c.number)
            row += 1
        data_final_row = row

        # creates the bar chart
        chart = workbook.add_chart({'type': 'bar'})
        chart.add_series({
            "name": "Number of 5-star tracks",
            'categories': "='# of 5-star tracks per artist'!A" + str(data_initial_row) + ":A" + str(data_final_row),
            'values':"='# of 5-star tracks per artist'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

        chart.set_y_axis({'name': 'Artist'})
        chart.set_x_axis({'name': 'Number of 5-star tracks'})
        worksheet.insert_chart('C1', chart, {'x_scale': 2, 'y_scale': 2})


        # number of 5-star tracks per album ============================================================================
        worksheet = workbook.add_worksheet("# of 5-star tracks per album")
        row = 1
        # writes the headers of the table
        worksheet.write("A" + str(row), "Album")
        worksheet.write("B" + str(row), "Number of 5-star tracks")
        row += 1

        counter = _counter.Counter()
        for artist in data.keys():
            for album in data[artist]:
                for title in data[artist][album]:
                    if int(data[artist][album][title]["rating"]) == 5:
                        counter.push(album)

        sorted_number = sorted(counter.entries, key=lambda value: value.number, reverse=True)

        data_initial_row = row
        for c in sorted_number:
            worksheet.write("A" + str(row), c.name)
            worksheet.write("B" + str(row), c.number)
            row += 1
        data_final_row = row

        # creates the bar chart
        chart = workbook.add_chart({'type': 'bar'})
        chart.add_series({
            "name": "Number of 5-star tracks",
            'categories': "='# of 5-star tracks per album'!A" + str(data_initial_row) + ":A" + str(data_final_row),
            'values':"='# of 5-star tracks per album'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

        chart.set_y_axis({'name': 'Album'})
        chart.set_x_axis({'name': 'Number of 5-star tracks'})
        worksheet.insert_chart('C1', chart, {'x_scale': 2, 'y_scale': 10})


        # # of musics per rating =======================================================================================
        worksheet = workbook.add_worksheet("# of musics per rating")
        row = 1
        # writes the headers of the table
        worksheet.write("A" + str(row), "Rating")
        worksheet.write("B" + str(row), "# of tracks")
        row += 1

        counter = _counter.Counter()
        for artist in data.keys():
            for album in data[artist]:
                for title in data[artist][album]:
                    counter.push(data[artist][album][title]["rating"])

        sorted_number = sorted(counter.entries, key=lambda value: value.name, reverse=True)

        data_initial_row = row
        for c in sorted_number:
            worksheet.write("A" + str(row), c.name)
            worksheet.write("B" + str(row), c.number)
            row += 1
        data_final_row = row

        # creates the bar chart
        chart = workbook.add_chart({'type': 'bar'})
        chart.add_series({
            "name": "Number of tracks per rating",
            'categories': "='# of musics per rating'!A" + str(data_initial_row) + ":A" + str(data_final_row),
            'values':"='# of musics per rating'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

        chart.set_y_axis({'name': 'Rating'})
        chart.set_x_axis({'name': 'Number of tracks'})
        worksheet.insert_chart('C1', chart, {'x_scale': 1, 'y_scale': 1})


        # # of tracks per genre ========================================================================================
        worksheet = workbook.add_worksheet("# of tracks per genre")
        row = 1
        # writes the headers of the table
        worksheet.write("A" + str(row), "Genre")
        worksheet.write("B" + str(row), "# of tracks")
        row += 1

        counter = _counter.Counter()
        for artist in data.keys():
            for album in data[artist]:
                for title in data[artist][album]:
                    counter.push(data[artist][album][title]["genre"])

        sorted_number = sorted(counter.entries, key=lambda value: value.number, reverse=True)

        data_initial_row = row
        for c in sorted_number:
            worksheet.write("A" + str(row), c.name)
            worksheet.write("B" + str(row), c.number)
            row += 1
        data_final_row = row

        # creates the bar chart
        chart = workbook.add_chart({'type': 'bar'})
        chart.add_series({
            "name": "Number of tracks per genre",
            'categories': "='# of tracks per genre'!A" + str(data_initial_row) + ":A" + str(data_final_row),
            'values':"='# of tracks per genre'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

        chart.set_y_axis({'name': 'Artist'})
        chart.set_x_axis({'name': 'Number of Tracks'})
        worksheet.insert_chart('C1', chart, {'x_scale': 2, 'y_scale': 2})


        # # of tracks per lenght range =================================================================================
        worksheet = workbook.add_worksheet("# of tracks per lenght range")
        row = 1
        worksheet.write("A" + str(row), "The lenght is inaccurate in songs with variable bit rate")
        row += 1

        # writes the headers of the table
        worksheet.write("A" + str(row), "Range")
        worksheet.write("B" + str(row), "Number of tracks")
        row += 1

        counter = _counter.Counter()
        for artist in data.keys():
            for album in data[artist]:
                for title in data[artist][album]:
                    if int(data[artist][album][title]["lenght"]) < 60: counter.push("0-1 minute")
                    elif int(data[artist][album][title]["lenght"]) < 120: counter.push("1-2 minutes")
                    elif int(data[artist][album][title]["lenght"]) < 180: counter.push("2-3 minutes")
                    elif int(data[artist][album][title]["lenght"]) < 240: counter.push("3-4 minutes")
                    elif int(data[artist][album][title]["lenght"]) < 300: counter.push("4-5 minutes")
                    elif int(data[artist][album][title]["lenght"]) < 360: counter.push("5-6 minutes")
                    elif int(data[artist][album][title]["lenght"]) < 420: counter.push("6-7 minutes")
                    elif int(data[artist][album][title]["lenght"]) < 480: counter.push("7-8 minutes")
                    elif int(data[artist][album][title]["lenght"]) >= 480: counter.push("more than 8 minutes")

        sorted_number = sorted(counter.entries, key=lambda value: value.name, reverse=True)

        data_initial_row = row
        for c in sorted_number:
            worksheet.write("A" + str(row), c.name)
            worksheet.write("B" + str(row), c.number)
            row += 1
        data_final_row = row

        # creates the bar chart
        chart = workbook.add_chart({'type': 'bar'})
        chart.add_series({
            "name": "Number of tracks per lenght range",
            'categories': "='# of tracks per lenght range'!A" + str(data_initial_row) + ":A" + str(data_final_row),
            'values':"='# of tracks per lenght range'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

        chart.set_y_axis({'name': 'Range'})
        chart.set_x_axis({'name': 'Number of Tracks'})
        worksheet.insert_chart('C1', chart, {'x_scale': 2, 'y_scale': 1})


        # avg rating per lenght range ==================================================================================
        worksheet = workbook.add_worksheet("avg rating per lenght range")
        row = 1
        worksheet.write("A" + str(row), "The lenght is inaccurate in songs with variable bit rate")
        row += 1

        # writes the headers of the table
        worksheet.write("A" + str(row), "Range")
        worksheet.write("B" + str(row), "Average Rating")
        row += 1

        rating_sum_counter = _counter.Counter()
        tracks_number_counter = _counter.Counter()
        for artist in data.keys():
            for album in data[artist]:
                for title in data[artist][album]:
                    if int(data[artist][album][title]["rating"]) > 0:
                        if int(data[artist][album][title]["lenght"]) < 60: tracks_number_counter.push("0-1 minute"); rating_sum_counter.push("0-1 minute", int(data[artist][album][title]["rating"]))
                        elif int(data[artist][album][title]["lenght"]) < 120: tracks_number_counter.push("1-2 minutes"); rating_sum_counter.push("1-2 minutes", int(data[artist][album][title]["rating"]))
                        elif int(data[artist][album][title]["lenght"]) < 180: tracks_number_counter.push("2-3 minutes"); rating_sum_counter.push("2-3 minutes", int(data[artist][album][title]["rating"]))
                        elif int(data[artist][album][title]["lenght"]) < 240: tracks_number_counter.push("3-4 minutes"); rating_sum_counter.push("3-4 minutes", int(data[artist][album][title]["rating"]))
                        elif int(data[artist][album][title]["lenght"]) < 300: tracks_number_counter.push("4-5 minutes"); rating_sum_counter.push("4-5 minutes", int(data[artist][album][title]["rating"]))
                        elif int(data[artist][album][title]["lenght"]) < 360: tracks_number_counter.push("5-6 minutes"); rating_sum_counter.push("5-6 minutes", int(data[artist][album][title]["rating"]))
                        elif int(data[artist][album][title]["lenght"]) < 420: tracks_number_counter.push("6-7 minutes"); rating_sum_counter.push("6-7 minutes", int(data[artist][album][title]["rating"]))
                        elif int(data[artist][album][title]["lenght"]) < 480: tracks_number_counter.push("7-8 minutes"); rating_sum_counter.push("7-8 minutes", int(data[artist][album][title]["rating"]))
                        elif int(data[artist][album][title]["lenght"]) >= 480: tracks_number_counter.push("more than 8 minutes"); rating_sum_counter.push("more than 8 minutes", int(data[artist][album][title]["rating"]))

        average_ratings = dict()
        for n in range(0, len(rating_sum_counter.entries)):
            average_ratings[rating_sum_counter.entries[n].name] = rating_sum_counter.entries[n].number / tracks_number_counter.entries[n].number if tracks_number_counter.entries[n].number > 0 else 0

        # gets the sorted list
        sorted_average_ratings = sorted(average_ratings.items(), key=lambda value: value[0], reverse=True)

        data_initial_row = row
        for a in sorted_average_ratings:
            worksheet.write("A" + str(row), a[0])
            worksheet.write("B" + str(row), a[1])
            row += 1
        data_final_row = row

        # creates the bar chart
        chart = workbook.add_chart({'type': 'bar'})
        chart.add_series({
            "name": "Average rating per lenght range",
            'categories': "='avg rating per lenght range'!A" + str(data_initial_row) + ":A" + str(data_final_row),
            'values':"='avg rating per lenght range'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

        chart.set_y_axis({'name': 'Range'})
        chart.set_x_axis({'name': 'Average Rating'})
        worksheet.insert_chart('C1', chart, {'x_scale': 2, 'y_scale': 1})


        # average rating per genre =====================================================================================
        worksheet = workbook.add_worksheet("average rating per genre")
        row = 1
        # writes the headers of the table
        worksheet.write("A" + str(row), "Genre")
        worksheet.write("B" + str(row), "Average Rating")
        row += 1

        rating_sum_counter = _counter.Counter()
        tracks_number_counter = _counter.Counter()
        for artist in data.keys():
            for album in data[artist]:
                for title in data[artist][album]:
                    if int(data[artist][album][title]["rating"]) > 0:
                        rating_sum_counter.push(data[artist][album][title]["genre"], int(data[artist][album][title]["rating"]))
                        tracks_number_counter.push(data[artist][album][title]["genre"])

        average_ratings = dict()
        for n in range(0, len(rating_sum_counter.entries)):
            average_ratings[rating_sum_counter.entries[n].name] = rating_sum_counter.entries[n].number / tracks_number_counter.entries[n].number if tracks_number_counter.entries[n].number > 0 else 0

        # gets the sorted list
        sorted_average_ratings = sorted(average_ratings.items(), key=lambda value: value[1], reverse=True)

        data_initial_row = row
        for a in sorted_average_ratings:
            worksheet.write("A" + str(row), a[0])
            worksheet.write("B" + str(row), a[1])
            row += 1
        data_final_row = row

        # creates the bar chart
        chart = workbook.add_chart({'type': 'bar'})
        chart.add_series({
            "name": "Average rating per lenght range",
            'categories': "='average rating per genre'!A" + str(data_initial_row) + ":A" + str(data_final_row),
            'values':"='average rating per genre'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

        chart.set_y_axis({'name': 'Genre'})
        chart.set_x_axis({'name': 'Average Rating'})
        worksheet.insert_chart('C1', chart, {'x_scale': 2, 'y_scale': 2})

        # closes the file
        workbook.close()
        print("Done!")
