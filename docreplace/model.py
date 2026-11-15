from dataclasses import dataclass, field
from typing import Any


@dataclass
class Unit:
    key: str
    location: str
    before: str
    after: str
    count: int
    kind: str = 'text'
    selected: bool = True
    payload: Any = None


@dataclass
class ScanResult:
    path: str
    fingerprint: str
    units: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    error: str = ''
    options: dict = field(default_factory=dict)


@dataclass
class SaveResult:
    source: str
    status: str
    destination: str = ''
    backup: str = ''
    detail: str = ''
