import sqlite3
import os

import _tag

def save(target_path, cursor):
    print("Reading files...")
    # walks the received folder searching for mp3 files
    for (dirpath, dirnames, filenames) in os.walk(target_path):
        for f in filenames:
            if os.path.splitext(f)[1] == ".mp3":
                tag = _tag.Tag(os.path.join(dirpath, f))
                # if the file was loaded correctly, add it to the database
                if not tag.title == "<none>":
                    try:
                        cursor.executescript(
                            "insert or replace into Track (title, album, artist, genre, lenght, rating) values (" +
                            "\"" + tag.title + "\", " +
                            "\"" + tag.album + "\", " +
                            "\"" + tag.artist + "\", " +
                            "\"" + tag.genre + "\", " +
                            "" + str(tag.lenght) + ", " +
                            "" + str(tag.rating) + ");")
                    except:
                        print("Error: could not save " + tag.to_string())