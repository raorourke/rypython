import PySimpleGUI as sg
from typing import List, Any


class NewTab:
    def __init__(self, tab_name: str, tab_layout: List[List[Any]]):
        self.tab = sg.Tab(
            tab_name,
            tab_layout
        )

    @property
    def config(self):
        return [
            self.tab
        ]


class NewTabGroup:
    def __init__(self, tabs: List[NewTab]):
        self.tab_group = sg.TabGroup(
            [
                tab.config
                for tab in tabs
            ]
        )

    @property
    def config(self):
        return [
            self.tab_group
        ]



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



