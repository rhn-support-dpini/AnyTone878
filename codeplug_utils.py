"""Shared utilities for AnyTone codeplug Excel/CSV processing."""

from __future__ import annotations

import csv
import re
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

LOG_FILE = "process.log"

CHANNEL_TYPE_ANALOG = "A-Analog"
CHANNEL_TYPE_DIGITAL = "D-Digital"
VALID_CHANNEL_TYPES = {CHANNEL_TYPE_ANALOG, CHANNEL_TYPE_DIGITAL}

COLOR_ANALOG = "E2EFDA"
COLOR_DIGITAL = "D9E1F2"

FREQUENCY_COLUMN_KEYWORDS = ("frequency",)


def is_single_frequency_column(column_name: str) -> bool:
    """True for frequency columns holding one value, not pipe-separated lists."""
    if not is_frequency_column(column_name):
        return False
    lowered = column_name.lower()
    if "member" in lowered:
        return False
    return True


def today_str() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def log_message(message: str, level: str = "INFO") -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {level}: {message}\n"
    with open(LOG_FILE, "a", encoding="utf-8") as fh:
        fh.write(line)
    print(line.rstrip())


def find_input_file(directory: Path, base_name: str) -> Path:
    """Locate a CSV input file regardless of extension case."""
    stem = Path(base_name).stem
    candidates = [
        directory / base_name,
        directory / base_name.lower(),
        directory / base_name.upper(),
        directory / f"{stem}.csv",
        directory / f"{stem}.CSV",
        directory / f"{stem.lower()}.csv",
        directory / f"{stem.upper()}.CSV",
    ]
    seen: set[Path] = set()
    for path in candidates:
        if path in seen:
            continue
        seen.add(path)
        if path.is_file():
            return path
    raise FileNotFoundError(f"File non trovato: {base_name} in {directory}")


def read_csv_dicts(path: Path) -> Tuple[List[str], List[Dict[str, str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        if reader.fieldnames is None:
            return [], []
        fieldnames = [name.strip() for name in reader.fieldnames]
        rows: List[Dict[str, str]] = []
        for raw in reader:
            row = {fieldnames[i]: (raw[fieldnames[i]] if fieldnames[i] in raw else "")
                   for i in range(len(fieldnames))}
            rows.append(row)
        return fieldnames, rows


def write_csv_dicts(path: Path, fieldnames: Sequence[str], rows: Sequence[Dict[str, str]]) -> None:
    with path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames, quoting=csv.QUOTE_ALL)
        writer.writeheader()
        for row in rows:
            writer.writerow({col: row.get(col, "") for col in fieldnames})


def is_frequency_column(column_name: str) -> bool:
    lowered = column_name.lower()
    return any(keyword in lowered for keyword in FREQUENCY_COLUMN_KEYWORDS)


def normalize_frequency(value: Any, context: str = "") -> str:
    """Format a frequency value as a string with exactly 5 decimal places."""
    if value is None:
        return ""

    text = str(value).strip()
    if text == "":
        return ""

    original = text
    text = text.replace(",", ".")

    if not re.fullmatch(r"-?\d+(?:\.\d+)?", text):
        log_message(
            f"Discrepanza frequenza: valore non valido '{original}' ({context}); "
            "impossibile normalizzare",
            "WARNING",
        )
        return original

    try:
        number = Decimal(text)
    except InvalidOperation:
        log_message(
            f"Discrepanza frequenza: valore non valido '{original}' ({context})",
            "WARNING",
        )
        return original

    quantized = number.quantize(Decimal("0.00001"), rounding=ROUND_HALF_UP)
    formatted = f"{quantized:.5f}"

    if "." in original:
        decimal_part = original.split(".", 1)[1]
        if len(decimal_part) != 5:
            log_message(
                f"Discrepanza frequenza: '{original}' riformattato a '{formatted}' ({context})",
                "WARNING",
            )
    elif formatted != f"{int(number)}.00000":
        log_message(
            f"Discrepanza frequenza: '{original}' riformattato a '{formatted}' ({context})",
            "WARNING",
        )

    return formatted


def normalize_frequencies_in_row(
    row: Dict[str, str],
    columns: Iterable[str],
    context: str = "",
) -> Dict[str, str]:
    updated = dict(row)
    for column in columns:
        if column in updated and updated[column].strip():
            updated[column] = normalize_frequency(updated[column], f"{context}/{column}")
    return updated


def split_pipe_values(value: str) -> List[str]:
    if not value or not str(value).strip():
        return []
    return [part.strip() for part in str(value).split("|") if part.strip()]


def join_pipe_values(values: Sequence[str]) -> str:
    return "|".join(values)


def is_numeric_index(value: str) -> bool:
    return bool(re.fullmatch(r"\d+", value.strip()))


def resolve_member_to_name(
    member: str,
    old_index_to_name: Dict[int, str],
    context: str,
) -> str:
    member = member.strip()
    if not member:
        return member
    if is_numeric_index(member):
        index = int(member)
        if index in old_index_to_name:
            return old_index_to_name[index]
        log_message(
            f"Riferimento mancante: indice {index} non trovato in Channel.csv ({context})",
            "WARNING",
        )
        return member
    return member


def resolve_members_field(
    value: str,
    old_index_to_name: Dict[int, str],
    context: str,
) -> str:
    if not value.strip():
        return value
    resolved = [
        resolve_member_to_name(member, old_index_to_name, context)
        for member in split_pipe_values(value)
    ]
    return join_pipe_values(resolved)


def ensure_unique_channel_names(
    channels: List[Dict[str, str]],
) -> Tuple[List[Dict[str, str]], Dict[int, str]]:
    """Assign unique channel names and build old_index -> name map."""
    old_index_to_name: Dict[int, str] = {}
    seen_names: Dict[str, int] = {}

    for position, row in enumerate(channels, start=1):
        name = (row.get("Channel Name") or "").strip()
        if not name:
            name = f"Channel_{position}"
            log_message(
                f"Nome canale vuoto alla riga {position}: assegnato '{name}'",
                "WARNING",
            )

        base_name = name
        if base_name in seen_names:
            seen_names[base_name] += 1
            name = f"{base_name}_{seen_names[base_name]}"
            log_message(
                f"Nome canale duplicato '{base_name}' alla riga {position}: "
                f"rinominato in '{name}'",
                "WARNING",
            )
        else:
            seen_names[base_name] = 0

        row["Channel Name"] = name
        old_index_to_name[position] = name

    return channels, old_index_to_name


def is_blank_excel_row(row: Dict[str, Any], channel_columns: Sequence[str]) -> bool:
    meaningful = [
        row.get("Channel Name", ""),
        row.get("Channel Type", ""),
        row.get("Receive Frequency", ""),
        row.get("Transmit Frequency", ""),
    ]
    return not any(str(v).strip() for v in meaningful)


def default_blank_channel_row(template: Dict[str, str]) -> Dict[str, str]:
    """Create a placeholder channel row preserving column structure."""
    blank = {key: "" for key in template}
    for key, value in template.items():
        if value in ("Off", "0", "1", "None", "High", "12.5K", "Carrier"):
            blank[key] = value
    blank["Channel Name"] = ""
    blank["Channel Type"] = ""
    blank["Scan List"] = "None"
    blank["Receive Group List"] = "None"
    return blank


def find_latest_master_excel(directory: Path) -> Optional[Path]:
    matches = sorted(directory.glob("Master_Codeplug_*.xlsx"), reverse=True)
    return matches[0] if matches else None
