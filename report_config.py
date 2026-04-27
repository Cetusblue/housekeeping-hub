from functools import lru_cache
from master_loader import (
    load_report_line_master_rows,
    load_report_mapping_rows,
)


@lru_cache(maxsize=1)
def _cached_report_lines():
    return load_report_line_master_rows()


@lru_cache(maxsize=1)
def _cached_report_mappings():
    return load_report_mapping_rows()


def get_report_lines():
    return _cached_report_lines()


def get_report_line_id_for_item(item_name: str):
    report_lines = _cached_report_lines()
    mappings = _cached_report_mappings()

    line_name_to_id = {
        r["report_line_name"]: r["report_line_id"]
        for r in report_lines
    }

    item_to_line_name = {
        m["item_name"]: m["report_line_name"]
        for m in mappings
    }

    report_line_name = item_to_line_name.get(item_name)
    if not report_line_name:
        return None

    return line_name_to_id.get(report_line_name)