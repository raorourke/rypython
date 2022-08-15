import PySimpleGUI as sg


class NewWindow:
    def __init__(self, name, layout):
        self.event = None
        self.values = None
        self.name = name
        self.layout = layout

    def __enter__(self):
        self.window = sg.Window(self.name, self.layout)
        return self

    def __exit__(self, type, value, traceback):
        self.window.close()

    def read(self):
        self.event, self.values = self.window.read()
