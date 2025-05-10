from dataclasses import dataclass
from typing import Optional


@dataclass
class Event:
    name: str
    department: str
    date: str
    participants: str
    responsible: str
    # ... остальные поля ...

    def to_dict(self) -> dict:
        return self.__dict__
