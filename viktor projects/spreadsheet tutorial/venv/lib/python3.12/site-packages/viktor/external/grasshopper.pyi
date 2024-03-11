from ..core import File as File
from .external_program import ExternalProgram as ExternalProgram
from _typeshed import Incomplete
from typing import Any, Dict

class _DataTree:
    data: Incomplete
    def __init__(self, name: str) -> None: ...
    def append(self, path: list, items: list) -> None: ...

class GrasshopperAnalysis(ExternalProgram):
    script: Incomplete
    def __init__(self, *, script: File, input_parameters: Dict[str, Any] = ...) -> None: ...
    def get_output(self) -> dict: ...
