from typing import List, Callable

import pendulum as pm
import pymsteams as teams


class TeamsAlert:
    def __init__(self, alert_title: str, bot_link: str):
        self.title = alert_title
        self.bot_link = bot_link

    def __call__(self, func):

        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                message = teams.connectorcard(self.bot_link)
                message.title(self.title)
                message.text(f"{e} ({pm.now():%Y-%m-%d %H:%M:%S})")
                message.send()

        return wrapper


class AdaptiveTextBlock:
    def __init__(
            self,
            text: str,
            wrap: bool = True,
            weight: str = 'Bolder'
    ):
        self.text = text
        self.wrap = wrap
        self.weight = weight

    @property
    def json(self):
        return {
            "type": "TextBlock",
            "text": self.text,
            "wrap": self.wrap,
            "weight": self.weight
        }


class AdaptiveContainer:
    def __init__(self, items: List[dict]):
        self.items = items

    @property
    def json(self):
        return {
            "type": "Container",
            "items": self.items,
            "separator": True
        }


class AdaptiveBulletList:
    def __init__(
            self,
            items: List[str],
            header: AdaptiveTextBlock = None,
            tfunc: Callable[[str], str] = None,
            wrap: bool = True,
            spacing: str = 'Small',
            weight: str = 'Default'
    ):
        self.header = header
        self.items = [
            tfunc(item)
            for item in items
        ] if tfunc is not None else items
        self.wrap = wrap
        self.spacing = spacing
        self.weight = weight

    @property
    def item_json(self):
        return {
            "type": "TextBlock",
            "text": ' \r'.join(
                f"- {item}"
                for item in self.items
            ),
            "wrap": self.wrap,
            "spacing": self.spacing,
            "weight": self.weight
        }

    @property
    def json(self):
        items = [
            self.header.json,
            self.item_json
        ] if self.header is not None else [
            self.item_json
        ]
        container = AdaptiveContainer(items)
        return container.json


class AdaptiveCard:
    def __init__(
            self,
            title: str,
            size: str = 'ExtraLarge',
            weight: str = 'Bolder'
    ):
        self.title = title
        self.size = size
        self.weight = weight
        self.body = [
            AdaptiveTextBlock
        ]

    @property
    def json(self):
        return {
            "type": "AdaptiveCard",
            "body": [
                {
                    "type": "TextBlock",
                    "size": self.size,
                    "weight": self.weight,
                    "text": self.title
                },
                *[
                    element.json
                    for element in self.body
                ]
            ],
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.0"
        }

    def add_text_block(self):
