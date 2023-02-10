from __future__ import annotations

import pandas as pd
from typing import Dict, List, Any, Union
from dataclasses import dataclass

class Mermaid:
    OPEN: str = "```mermaid"
    CLOSE: str = "```"
    INDENT: str = "    "

    def __init__(self, declaration: str):
        self.declaration = declaration
        self.lines = []

    def add_line(self, new_line: str):
        self.lines.append(new_line)

    def print(self):
        lines = [
            self.OPEN,
            self.declaration,
            *self.lines,
            self.CLOSE
        ]
        print("\n".join(lines))


@dataclass
class Section:
    name: str

    @property
    def declaration(self):
        return f"    section {self.name}"


@dataclass
class Task:
    name: str
    description: str
    turnaround: str
    is_critical: bool | str = False
    after: List[str] = None
    start_date: str = ""
    end_date: str = ""

    @property
    def is_active(self):
        return (bool(self.start_date) and not bool(self.end_date))

    @property
    def is_milestone(self):
        if self.name.lower() == "milestone":
            return True
        return False

    @property
    def is_done(self):
        if self.is_milestone:
            return False
        return bool(self.end_date)

    @property
    def declaration(self):
        if self.is_milestone:
            return f"    {self.description}\t:milestone, {self.turnaround}, 0d"
        crit = "crit, " if self.is_critical else ""
        status = ""
        if self.is_active:
            status = "active, "
        if self.is_done:
            status = "done, "
        after = f" after {' '.join(self.after)}," if self.after is not None else ""
        end_date = self.end_date or self.turnaround
        return f"    {self.description}\t:{crit}{status}{self.name},{after} {self.start_date}, {end_date}"




class ganttChart(Mermaid):
    def __init__(
            self,
            title: str,
            end_date: str,
            date_format: str = "YYYY-MM-DD",
            excludes: List[str] = None,
            note: str = None
    ):
        super().__init__(declaration="gantt")
        self.title = title
        self.end_date = end_date
        self.date_format = date_format
        self.excludes = ", ".join(excludes) if excludes else ""
        self.note = note
        self._add_description_lines()

    def _add_description_lines(self):
        date_format = f"{self.INDENT}dateFormat\t{self.date_format}"
        self.lines.append(date_format)
        title = f"{self.INDENT}title\t{self.title}"
        self.lines.append(title)
        excludes = f"{self.INDENT}excludes\t{self.excludes}"
        self.lines.append(excludes)
        if self.note is not None:
            note = f"{self.INDENT}%% {self.note}"
            self.lines.append(note)

    def add_section(
            self,
            name: str,
            tasks: List[Task]
    ):
        self.lines.append(self.INDENT)
        section = Section(name=name)
        self.lines.append(section.declaration)
        for task in tasks:
            self.lines.append(task.declaration)






class classDiagram(Mermaid):
    def __init__(self):
        super().__init__(declaration="classDiagram")

    def add_class(
            self,
            class_instance: Any,
            ignore_private: bool = True
    ):
        class_name = class_instance.__class__.__name__
        self.lines.append(f"{self.INDENT} class {class_name} {{")
        attr_dict = {}
        sorted_attr = {}
        all_attrs = dir(class_instance)
        for attr_name in all_attrs:
            try:
                attr_dict[attr_name] = getattr(class_instance, attr_name)
            except AttributeError:
                continue

        for attr_name, attr in attr_dict.items():
            if attr_name.startswith("_") and ignore_private:
                continue
            if attr_name in class_instance.__dict__:
                sorted_attr.setdefault(
                    "attributes", []
                ).append((attr_name, type(attr).__name__))
                continue
            if f"bound method {class_name}" in repr(attr):
                sorted_attr.setdefault(
                    "methods", []
                ).append(attr_name)

        attributes = sorted(sorted_attr.get("attributes"))
        for attr_name, attr_type in attributes:
            self.lines.append(f"{self.INDENT*2} +{attr_type} {attr_name}")

        methods = sorted(sorted_attr.get("methods"))
        for method_name in methods:
            self.lines.append(f"{self.INDENT * 2} +{method_name}()")
        self.add_line(f"{self.INDENT}}}")







class erDiagram(Mermaid):
    def __init__(self):
        super().__init__(declaration="erDiagram")

    def add_entity(
            self,
            entity_name: str,
            dtypes: Dict[str, str],
            primary_key: str = None,
            foreign_keys: List[str] = None,
            descriptions: Dict[str, str] = None
    ):
        foreign_keys = foreign_keys or []
        descriptions = descriptions or {}
        self.add_line(f"{self.INDENT}{entity_name} {{")
        for column_name, dtype in dtypes.items():
            base_line = f"{self.INDENT*2}{dtype} {column_name}"
            key_status = ""
            if column_name == primary_key:
                key_status = "PK"
            if column_name in foreign_keys:
                key_status = "FK"
            description = descriptions.get(column_name, "")
            full_line = f"{base_line} {key_status} {description}"
            self.add_line(full_line)
        self.add_line(f"{self.INDENT}}}")



    def add_entity_by_dtypes(
            self,
            entity_name: str,
            dtypes: pd.Series,
            primary_key: str = None,
            foreign_keys: List[str] = None,
            descriptions: Dict[str, str] = None
    ):
        dtypes = {
            column_name: dtype.name
            for column_name, dtype in dtypes.copy().reset_index().values.tolist()
        }
        self.add_entity(
            entity_name,
            dtypes,
            primary_key=primary_key,
            foreign_keys=foreign_keys,
            descriptions=descriptions
        )

