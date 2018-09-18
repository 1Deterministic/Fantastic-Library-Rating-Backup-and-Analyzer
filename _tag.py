import math

import eyed3
from eyed3.id3 import ID3_V1_0, ID3_V1_1, ID3_V2_3, ID3_V2_4
#from mutagen.mp3 import MP3

class Tag:
    # creates a tag object containing some information
    def __init__(self, path):
        # tries to read information from the file
        try:
            self.file = eyed3.load(path)

            # individual attributes can fail too

            try: self.artist = self.file.tag.artist
            except: self.artist = "<none>"

            try: self.album = self.file.tag.album
            except: self.album = "<none>"

            try: self.title = self.file.tag.title
            except: self.title = "<none>"

            try: self.genre = self.file.tag.genre.name
            except: self.genre = "<none>"

            try: self.lenght = self.file.info.time_secs  # MP3("path").info.length
            except: self.lenght = 0

            try: self.rating = math.ceil(self.file.tag.frame_set[b'POPM'][0].rating / 51)
            except: self.rating = 0

        except:
            # erases all properties
            self.artist = "<none>"
            self.album = "<none>"
            self.title = "<none>"
            self.genre = "<none>"
            self.lenght = 0
            self.rating = 0



    # writes the current properties to the file
    def write(self):
        # removes a known problematic frame
        try: del(self.file.tag.frame_set[b'RGAD'])
        except: pass

        if self.artist != "<none>": self.file.tag.artist = self.artist
        if self.album != "<none>": self.file.tag.album = self.album
        if self.title != "<none>": self.file.tag.title = self.title
        if self.genre != "<none>": self.file.tag.genre = self.genre
        self.file.tag.frame_set[b'POPM'][0].rating = self.rating * 51

        self.file.tag.save()
        # change to return correctly
        return True

    # creates a string describing this tag
    def to_string(self):
        return self.artist + ":" + self.album + ":" + self.title