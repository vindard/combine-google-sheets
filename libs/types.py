from typing import List, Dict, Sequence, Tuple, Union, Any

from gspread.models import Spreadsheet, Worksheet


# Custom types
# TODO fix SheetDict to stop throwing assignment mypy errors below
SheetDict = Dict[str, Union[int, str, dict, Sequence, Worksheet]]
SheetDicts = Dict[str, SheetDict]
Row = List[str]
SheetValues = List[Row]
