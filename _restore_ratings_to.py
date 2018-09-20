import sqlite3
import os

import _tag

from fuzzywuzzy import fuzz

def restore(target_path, cursor, approximation):
    artist_min_approximation_ratio = 95 # artist names usually doesn't vary too much
    album_min_approximation_ratio = 85 # album title varies more
    title_min_approximation_ratio = 70 # track title varies a lot, so it is more permissive

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
            if (approximation and (fuzz.token_set_ratio(line[0], f.title) > title_min_approximation_ratio and
                                   fuzz.token_set_ratio(line[1], f.album) > album_min_approximation_ratio and
                                   fuzz.token_set_ratio(line[2], f.artist) > artist_min_approximation_ratio)) \
                or (line[0] == f.title and
                    line[1] == f.album and
                    line[2] == f.artist):

                f.rating = line[5];
                if not f.write():
                    print("Error: could not restore rating to " + f.to_string())
                break
    print("Done!")