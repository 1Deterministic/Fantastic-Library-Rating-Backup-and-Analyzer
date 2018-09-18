class Counter:
    def __init__(self):
        self.entries = []

    def push(self, key, value = None):
        if value is None:
            value = 1

        for entry in self.entries:
            if entry.name == key:
                entry.sum(value)
                return

        self.entries.append(Entry(key, value))



class Entry:
    def __init__(self, name, value):
        self.name = name
        self.number = value

    def sum(self, value = None):
        self.number += value
