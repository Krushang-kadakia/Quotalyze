# src/schema.py
from dataclasses import dataclass
from typing import List

@dataclass(frozen=True)
class QuotationSchema:
    all_columns: List[str]
    master_columns: List[str]
    vendor_value_columns: List[str]
    header_row_index: int = 0
