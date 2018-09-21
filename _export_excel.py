import collections
import sqlite3

import xlsxwriter
from fuzzywuzzy import fuzz

import _counter

def export(cursor):
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

    # writes the xlsx file
    print("Writing to .xlsx file...")
    workbook = xlsxwriter.Workbook("analytics.xlsx")


    # all information ==============================================================================================
    # adds a new sheet to the file
    page_name = "all tracks info"
    worksheet = workbook.add_worksheet(page_name)
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
    page_name = "average track rating per artist"
    worksheet = workbook.add_worksheet(page_name)
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
        average_ratings[rating_sum_counter.entries[n].name] = rating_sum_counter.entries[n].number / \
                                                              tracks_number_counter.entries[n].number if \
        tracks_number_counter.entries[n].number > 0 else 0

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
        'categories': "='" + page_name + "'!A" + str(data_initial_row) + ":A" + str(data_final_row),
        'values': "='" + page_name +"'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

    chart.set_y_axis({'name': 'Artist'})
    chart.set_x_axis({'name': 'Average Rating'})
    worksheet.insert_chart('C1', chart, {'x_scale': 2, 'y_scale': 5})


    # average track rating per album ===============================================================================
    page_name = "average track rating per album"
    worksheet = workbook.add_worksheet(page_name)
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
        average_ratings[rating_sum_counter.entries[n].name] = rating_sum_counter.entries[n].number / \
                                                              tracks_number_counter.entries[n].number if \
        tracks_number_counter.entries[n].number > 0 else 0

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
        'categories': "='" + page_name + "'!A" + str(data_initial_row) + ":A" + str(data_final_row),
        'values': "='" + page_name + "'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

    chart.set_y_axis({'name': 'Album'})
    chart.set_x_axis({'name': 'Average Rating'})
    worksheet.insert_chart('C1', chart, {'x_scale': 2, 'y_scale': 2})


    # number of 5-star tracks per artist ===========================================================================
    page_name = "# of 5-star tracks per artist"
    worksheet = workbook.add_worksheet(page_name)
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
        'categories': "='" + page_name + "'!A" + str(data_initial_row) + ":A" + str(data_final_row),
        'values': "='" + page_name + "'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

    chart.set_y_axis({'name': 'Artist'})
    chart.set_x_axis({'name': 'Number of 5-star tracks'})
    worksheet.insert_chart('C1', chart, {'x_scale': 2, 'y_scale': 2})


    # number of 5 and 4-star tracks per artist =========================================================================
    page_name = "# of >=4-star tracks per artist"
    worksheet = workbook.add_worksheet(page_name)
    row = 1
    # writes the headers of the table
    worksheet.write("A" + str(row), "Artist")
    worksheet.write("B" + str(row), "Number of 5 or 4-star tracks")
    row += 1

    counter = _counter.Counter()
    for artist in data.keys():
        for album in data[artist]:
            for title in data[artist][album]:
                if int(data[artist][album][title]["rating"]) >= 4:
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
        "name": "Number of 5 or 4-star tracks",
        'categories': "='" + page_name + "'!A" + str(data_initial_row) + ":A" + str(data_final_row),
        'values': "='" + page_name + "'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

    chart.set_y_axis({'name': 'Artist'})
    chart.set_x_axis({'name': 'Number of 5-star tracks'})
    worksheet.insert_chart('C1', chart, {'x_scale': 2, 'y_scale': 2})


    # number of 5-star tracks per album ============================================================================
    page_name = "# of 5-star tracks per album"
    worksheet = workbook.add_worksheet(page_name)
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
        'categories': "='" + page_name + "'!A" + str(data_initial_row) + ":A" + str(data_final_row),
        'values': "='" + page_name + "'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

    chart.set_y_axis({'name': 'Album'})
    chart.set_x_axis({'name': 'Number of 5-star tracks'})
    worksheet.insert_chart('C1', chart, {'x_scale': 2, 'y_scale': 10})


    # # of musics per rating =======================================================================================
    page_name = "# of musics per rating"
    worksheet = workbook.add_worksheet(page_name)
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
        'categories': "='" + page_name + "'!A" + str(data_initial_row) + ":A" + str(data_final_row),
        'values': "='" + page_name + "'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

    chart.set_y_axis({'name': 'Rating'})
    chart.set_x_axis({'name': 'Number of tracks'})
    worksheet.insert_chart('C1', chart, {'x_scale': 1, 'y_scale': 1})


    # # of tracks per genre ========================================================================================
    page_name = "# of tracks per genre"
    worksheet = workbook.add_worksheet(page_name)
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
        'categories': "='" + page_name + "'!A" + str(data_initial_row) + ":A" + str(data_final_row),
        'values': "='" + page_name + "'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

    chart.set_y_axis({'name': 'Artist'})
    chart.set_x_axis({'name': 'Number of Tracks'})
    worksheet.insert_chart('C1', chart, {'x_scale': 2, 'y_scale': 2})


    # # of tracks per lenght range =================================================================================
    page_name = "# of tracks per lenght range"
    worksheet = workbook.add_worksheet(page_name)
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
                if int(data[artist][album][title]["lenght"]) < 60:
                    counter.push("0-1 minute")
                elif int(data[artist][album][title]["lenght"]) < 120:
                    counter.push("1-2 minutes")
                elif int(data[artist][album][title]["lenght"]) < 180:
                    counter.push("2-3 minutes")
                elif int(data[artist][album][title]["lenght"]) < 240:
                    counter.push("3-4 minutes")
                elif int(data[artist][album][title]["lenght"]) < 300:
                    counter.push("4-5 minutes")
                elif int(data[artist][album][title]["lenght"]) < 360:
                    counter.push("5-6 minutes")
                elif int(data[artist][album][title]["lenght"]) < 420:
                    counter.push("6-7 minutes")
                elif int(data[artist][album][title]["lenght"]) < 480:
                    counter.push("7-8 minutes")
                elif int(data[artist][album][title]["lenght"]) >= 480:
                    counter.push("more than 8 minutes")

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
        'categories': "='" + page_name + "'!A" + str(data_initial_row) + ":A" + str(data_final_row),
        'values': "='" + page_name + "'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

    chart.set_y_axis({'name': 'Range'})
    chart.set_x_axis({'name': 'Number of Tracks'})
    worksheet.insert_chart('C1', chart, {'x_scale': 2, 'y_scale': 1})


    # avg rating per lenght range ==================================================================================
    page_name = "avg rating per lenght range"
    worksheet = workbook.add_worksheet(page_name)
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
                    if int(data[artist][album][title]["lenght"]) < 60:
                        tracks_number_counter.push("0-1 minute"); rating_sum_counter.push("0-1 minute", int(
                            data[artist][album][title]["rating"]))
                    elif int(data[artist][album][title]["lenght"]) < 120:
                        tracks_number_counter.push("1-2 minutes"); rating_sum_counter.push("1-2 minutes", int(
                            data[artist][album][title]["rating"]))
                    elif int(data[artist][album][title]["lenght"]) < 180:
                        tracks_number_counter.push("2-3 minutes"); rating_sum_counter.push("2-3 minutes", int(
                            data[artist][album][title]["rating"]))
                    elif int(data[artist][album][title]["lenght"]) < 240:
                        tracks_number_counter.push("3-4 minutes"); rating_sum_counter.push("3-4 minutes", int(
                            data[artist][album][title]["rating"]))
                    elif int(data[artist][album][title]["lenght"]) < 300:
                        tracks_number_counter.push("4-5 minutes"); rating_sum_counter.push("4-5 minutes", int(
                            data[artist][album][title]["rating"]))
                    elif int(data[artist][album][title]["lenght"]) < 360:
                        tracks_number_counter.push("5-6 minutes"); rating_sum_counter.push("5-6 minutes", int(
                            data[artist][album][title]["rating"]))
                    elif int(data[artist][album][title]["lenght"]) < 420:
                        tracks_number_counter.push("6-7 minutes"); rating_sum_counter.push("6-7 minutes", int(
                            data[artist][album][title]["rating"]))
                    elif int(data[artist][album][title]["lenght"]) < 480:
                        tracks_number_counter.push("7-8 minutes"); rating_sum_counter.push("7-8 minutes", int(
                            data[artist][album][title]["rating"]))
                    elif int(data[artist][album][title]["lenght"]) >= 480:
                        tracks_number_counter.push("more than 8 minutes"); rating_sum_counter.push(
                            "more than 8 minutes", int(data[artist][album][title]["rating"]))

    average_ratings = dict()
    for n in range(0, len(rating_sum_counter.entries)):
        average_ratings[rating_sum_counter.entries[n].name] = rating_sum_counter.entries[n].number / \
                                                              tracks_number_counter.entries[n].number if \
        tracks_number_counter.entries[n].number > 0 else 0

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
        'categories': "='" + page_name + "'!A" + str(data_initial_row) + ":A" + str(data_final_row),
        'values': "='" + page_name + "'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

    chart.set_y_axis({'name': 'Range'})
    chart.set_x_axis({'name': 'Average Rating'})
    worksheet.insert_chart('C1', chart, {'x_scale': 2, 'y_scale': 1})


    # average rating per genre =====================================================================================
    page_name = "average rating per genre"
    worksheet = workbook.add_worksheet(page_name)
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
                    rating_sum_counter.push(data[artist][album][title]["genre"],
                                            int(data[artist][album][title]["rating"]))
                    tracks_number_counter.push(data[artist][album][title]["genre"])

    average_ratings = dict()
    for n in range(0, len(rating_sum_counter.entries)):
        average_ratings[rating_sum_counter.entries[n].name] = rating_sum_counter.entries[n].number / \
                                                              tracks_number_counter.entries[n].number if \
        tracks_number_counter.entries[n].number > 0 else 0

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
        "name": "Average rating per genre",
        'categories': "='" + page_name + "'!A" + str(data_initial_row) + ":A" + str(data_final_row),
        'values': "='" + page_name + "'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

    chart.set_y_axis({'name': 'Genre'})
    chart.set_x_axis({'name': 'Average Rating'})
    worksheet.insert_chart('C1', chart, {'x_scale': 2, 'y_scale': 2})

    # number of 5-star tracks per genre ============================================================================
    page_name = "# of 5-star tracks per genre"
    worksheet = workbook.add_worksheet(page_name)
    row = 1
    # writes the headers of the table
    worksheet.write("A" + str(row), "Genre")
    worksheet.write("B" + str(row), "Number of 5-star tracks")
    row += 1

    counter = _counter.Counter()
    for artist in data.keys():
        for album in data[artist]:
            for title in data[artist][album]:
                if int(data[artist][album][title]["rating"]) == 5:
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
        "name": "Number of 5-star tracks",
        'categories': "='" + page_name + "'!A" + str(data_initial_row) + ":A" + str(data_final_row),
        'values': "='" + page_name + "'!$B$" + str(data_initial_row) + ":$B$" + str(data_final_row)})

    chart.set_y_axis({'name': 'Genre'})
    chart.set_x_axis({'name': 'Number of 5-star tracks'})
    worksheet.insert_chart('C1', chart, {'x_scale': 1, 'y_scale': 3})

    # closes the file
    workbook.close()
    print("Done!")
