#!/usr/bin/env python3
"""
Generate per-table PostgreSQL INSERT/UPSERT scripts from an Excel workbook and a schema.txt (CREATE TABLE dump).

Key features (as configured by user requirements):
- One .sql file per table, with order prefix based on FK dependency sort and a UTC timestamp in the filename.
- Per-file transaction (BEGIN/COMMIT).
- Batched multi-row INSERT ... VALUES (...), (...), ... ON CONFLICT (...) DO UPDATE SET ...;
- Conflict target derived from PK (preferred) or first UNIQUE constraint.
- Update all non-key columns on conflict.
- Excel sheet names must equal table names (case-sensitive).
- Type coercion for booleans, numerics (normalize commas/thousand separators), timestamps (UTC), dates.
- Empty cells -> NULL except NOT NULL columns:
    * If schema DEFAULT exists -> use DEFAULT (by inserting the DEFAULT keyword in the VALUES tuple)
    * Else: for numeric -> 0 (or 0.0 for double precision), boolean -> FALSE, text -> '' (empty string).
    * For date/timestamp WITHOUT an explicit DEFAULT -> **row fails and is logged** (to avoid inventing epoch/zero dates).
- UUID v4 auto-generation when type uuid and cell missing.
- Validate varchar(n) lengths; exceed -> row fails and is logged.
- If unsupported column types (json/jsonb/bytea/arrays/enums) are encountered -> program aborts with message.
- A plain-text report.txt summarizes warnings/errors and decisions.
- Finds the Excel file by scanning OUTPUT_DIR for the first *.xlsx or *.xls file; schema file must be named schema.txt there.

Dependencies: pandas, openpyxl
Tested on Python 3.10+
"""

import os
from pathlib import Path
import re
import sys
import uuid
import math
import decimal
from dataclasses import dataclass, field
from datetime import datetime, timezone, date, time
from typing import Dict, List, Optional, Tuple, Any
from textwrap import dedent

import pandas as pd
import json

# ---------------------------
# Configuration
# ---------------------------

# IMPORTANT: Set this to the target folder where schema.txt and the Excel workbook reside,
# and where you want outputs written.
OUTPUT_DIR = Path(os.path.join(os.path.dirname(__file__), "data")).resolve()

# Batch size for multi-row INSERTs
BATCH_SIZE = 500

# SQL target Postgres version (used only for documentation)
PG_VERSION = "12"

# ---------------------------
# Models
# ---------------------------

@dataclass
class ColumnDef:
    name: str
    data_type: str
    nullable: bool
    default: Optional[str]  # raw SQL default expression, e.g., "now()", "gen_random_uuid()", "'foo'"


@dataclass
class ConstraintInfo:
    primary_key: List[str] = field(default_factory=lambda: [])  # pyright: ignore[reportUnknownVariableType] # column names
    unique_constraints: List[List[str]] = field(default_factory=lambda: [])  # list of unique column lists
    foreign_keys: List[Tuple[List[str], str, List[str]]] = field(default_factory=lambda: [])  # (local_cols, ref_table_qualified, ref_cols)


@dataclass
class TableDef:
    schema: str  # e.g., public
    name: str    # unquoted name as appears in schema (case-sensitive from dump)
    columns: Dict[str, ColumnDef]  # by column name (case-sensitive to match dump)
    constraints: ConstraintInfo


# ---------------------------
# Utilities
# ---------------------------

def now_utc_stamp_for_filename() -> str:
    # ISO-like but filename-safe (no colons)
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H-%M-%S")


def q_ident(ident: str) -> str:
    # Quote identifier with double quotes, escape any internal double quotes by doubling
    return '"' + ident.replace('"', '""') + '"'


def q_qualified(schema: str, table: str) -> str:
    return f'{q_ident(schema)}.{q_ident(table)}'


def sql_literal_text(val: str) -> str:
    # Escape single quotes by doubling them; escape newlines with \n inside the string
    val = val.replace("\\", "\\\\")
    val = val.replace("'", "''")
    val = val.replace("\r\n", "\n").replace("\r", "\n")
    return f"'{val}'"


def sql_literal_bool(b: bool) -> str:
    return "TRUE" if b else "FALSE"


def sql_literal_int(n: int) -> str:
    return str(n)


def sql_literal_float(x: float) -> str:
    if math.isnan(x) or math.isinf(x):
        # Not representable: fallback to NULL (shouldn't happen after validation)
        return "NULL"
    # Use plain repr; Postgres accepts standard decimal floats
    return format(x, ".15g")


def sql_literal_numeric(dec: decimal.Decimal) -> str:
    return str(dec)  # Decimal to string preserves scale


def sql_literal_date(d: date) -> str:
    return f"DATE '{d.isoformat()}'"


def sql_literal_timestamp_utc(dt: datetime) -> str:
    # Always output with timezone UTC offset
    # Postgres accepts 'YYYY-MM-DD HH:MM:SS+00'
    dt = dt.astimezone(timezone.utc)
    return f"TIMESTAMPTZ '{dt.strftime('%Y-%m-%d %H:%M:%S+00')}'"


def sql_literal_timestamp_naive(dt: datetime) -> str:
    # For timestamp without time zone; we still treat input as UTC but then drop tz info
    return f"TIMESTAMP '{dt.strftime('%Y-%m-%d %H:%M:%S')}'"


def parse_decimal_normalize(value: Any, scale: Optional[int]) -> Optional[decimal.Decimal]:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    if isinstance(value, (int, float, decimal.Decimal)):
        d = decimal.Decimal(str(value))
    else:
        s = str(value).strip()
        if s == "":
            return None
        # normalize thousand separators (spaces or commas), and decimal comma -> dot
        s = s.replace(" ", "").replace(",", ".")
        try:
            d = decimal.Decimal(s)
        except decimal.InvalidOperation:
            raise ValueError(f"Invalid numeric literal: {value}")
    if scale is not None:
        # Round to scale using quantize
        q = decimal.Decimal(1).scaleb(-scale)  # 10^-scale
        d = d.quantize(q, rounding=decimal.ROUND_HALF_UP)
    return d


def normalize_boolean(value: Any) -> Optional[bool]:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    if isinstance(value, bool):
        return value
    s = str(value).strip().lower()
    if s in ("1", "true", "t", "yes", "y"):
        return True
    if s in ("0", "false", "f", "no", "n"):
        return False
    raise ValueError(f"Unrecognized boolean literal: {value}")


# def parse_timestamp_utc(value: Any, assume_noon_if_date_only: bool=True) -> Optional[datetime]:
#     if value is None or (isinstance(value, float) and math.isnan(value)):
#         return None
#     if isinstance(value, datetime):
#         # Treat as UTC (Excel datetimes can be naive; we treat as UTC-kind per user)
#         dt = value
#         if dt.tzinfo is None:
#             dt = dt.replace(tzinfo=timezone.utc)
#         else:
#             dt = dt.astimezone(timezone.utc)
#         return dt
#     if isinstance(value, date) and not isinstance(value, datetime):
#         # Date only -> noon UTC if allowed
#         if assume_noon_if_date_only:
#             dt = datetime.combine(value, time(12, 0, 0, tzinfo=timezone.utc))
#             return dt
#         else:
#             return None
#     s = str(value).strip()
#     if s == "":
#         return None
#     # Try pandas to_datetime with utc True
#     try:
#         dt = pd.to_datetime(s, utc=True, errors="raise")
#         # If it parsed as date-only (00:00Z), and we want noon for date-only inputs:
#         # Heuristic: if original string lacks time separator, treat as date-only.
#         if assume_noon_if_date_only and ("T" not in s and " " not in s and ":" not in s):
#             dt = dt + pd.Timedelta(hours=12)
#         return dt.to_pydatetime()
#     except Exception:
#         raise ValueError(f"Invalid timestamp: {value}")

def parse_timestamp_utc(value: Any, assume_noon_if_date_only: bool=True) -> Optional[datetime]:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    if isinstance(value, datetime):
        dt = value
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        else:
            dt = dt.astimezone(timezone.utc)
        return dt
    if isinstance(value, date) and not isinstance(value, datetime):
        if assume_noon_if_date_only:
            dt = datetime.combine(value, time(12, 0, 0, tzinfo=timezone.utc))
            return dt
        else:
            return None
    # Handle Excel serial numbers
    try:
        # Try to parse as float if it's a string
        serial = None
        if isinstance(value, (int, float)) and value > 59:
            serial = float(value)
        elif isinstance(value, str):
            s = value.strip()
            if s and s.isdigit() and float(s) > 59:
                serial = float(s)
        if serial is not None:
            dt = pd.to_datetime('1899-12-30') + pd.Timedelta(days=serial)
            dt = dt.tz_localize('UTC')
            return dt.to_pydatetime()
    except Exception:
        pass
    s = str(value).strip()
    if s == "":
        return None
    try:
        dt = pd.to_datetime(s, utc=True, errors="raise")
        if assume_noon_if_date_only and ("T" not in s and " " not in s and ":" not in s):
            dt = dt + pd.Timedelta(hours=12)
        return dt.to_pydatetime()
    except Exception:
        raise ValueError(f"Invalid timestamp: {value}")

def parse_date_only(value: Any) -> Optional[date]:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    if isinstance(value, date):
        if isinstance(value, datetime):
            return value.date()
        return value
    s = str(value).strip()
    if s == "":
        return None
    try:
        d = pd.to_datetime(s, utc=False, errors="raise").date()
        return d
    except Exception:
        raise ValueError(f"Invalid date: {value}")


# ---------------------------
# Schema Parsing
# ---------------------------

CREATE_TABLE_RE = re.compile(
    r"CREATE\s+TABLE\s+(?P<qualified>(?:(?P<schema>\"?[A-Za-z_][A-Za-z0-9_]*\"?)\.)?\"?(?P<table>[A-Za-z_][A-Za-z0-9_]*)\"?)\s*\((?P<body>.*?)\)\s*;",
    re.IGNORECASE | re.DOTALL,
)

# COLUMN_DEF_RE = re.compile(
#     r"^\s*(?P<col>\"?[A-Za-z_][A-Za-z0-9_]*\"?)\s+"
#     r"(?P<type>[^\s,]+(?:\s*\([^)]+\))?(?:\s*\[[^\]]+\])?)"
#     r"(?:\s+COLLATE\s+[^\s,]+)?"
#     r"(?:\s+(?P<null>NOT\s+NULL|NULL))?"
#     r"(?:\s+DEFAULT\s+(?P<default>[^,]+?))?"
#     r"\s*(?:,)?$",
#     re.IGNORECASE,
# )

COLUMN_DEF_RE = re.compile(
    r"^\s*(?P<col>\"?[A-Za-z_][A-Za-z0-9_]*\"?)\s+"
    r"(?P<type>.+?)(?=(?:\s+(?:COLLATE|NOT\s+NULL|NULL|DEFAULT|CONSTRAINT|CHECK|REFERENCES|UNIQUE|PRIMARY\s+KEY))|,|$)"
    r"(?:\s+COLLATE\s+[^\s,]+)?"
    r"(?:\s+(?P<null>NOT\s+NULL|NULL))?"
    r"(?:\s+DEFAULT\s+(?P<default>[^,]+?))?"
    r"\s*(?:,)?$",
    re.IGNORECASE,
)

PK_RE = re.compile(
    r"CONSTRAINT\s+\"?[A-Za-z0-9_]+\"?\s+PRIMARY\s+KEY\s*\((?P<cols>[^)]+)\)",
    re.IGNORECASE,
)

UNIQUE_RE = re.compile(
    r"CONSTRAINT\s+\"?[A-Za-z0-9_]+\"?\s+UNIQUE\s*\((?P<cols>[^)]+)\)",
    re.IGNORECASE,
)

FK_RE = re.compile(
    r"CONSTRAINT\s+\"?[A-Za-z0-9_]+\"?\s+FOREIGN\s+KEY\s*\((?P<cols>[^)]+)\)\s+REFERENCES\s+(?P<reftab>\"?[A-Za-z_][A-Za-z0-9_]*\"?(?:\.\s*\"?[A-Za-z_][A-Za-z0-9_]*\"?)?)\s*\((?P<refcols>[^)]+)\)",
    re.IGNORECASE,
)

IDENT_LIST_SPLIT_RE = re.compile(r"\s*,\s*")


def strip_quotes(ident: str) -> str:
    s = ident.strip()
    if s.startswith('"') and s.endswith('"'):
        s = s[1:-1].replace('""', '"')
    return s


def parse_ident_list(text: str) -> List[str]:
    return [strip_quotes(x) for x in IDENT_LIST_SPLIT_RE.split(text.strip()) if x.strip()]


def parse_schema(schema_text: str) -> Dict[str, TableDef]:
    tables: Dict[str, TableDef] = {}
    for m in CREATE_TABLE_RE.finditer(schema_text):
        schema = m.group("schema")
        table = m.group("table")
        schema = strip_quotes(schema) if schema else "public"
        table = strip_quotes(table)

        body = m.group("body")
        # Split body into lines/chunks by commas that are line-separated; a robust approach is to split on commas not inside parentheses.
        parts: List[str] = []
        buf: List[str] = []
        depth = 0
        for ch in body:
            if ch == '(':
                depth += 1
            elif ch == ')':
                depth -= 1
            if ch == ',' and depth == 0:
                parts.append(''.join(buf).strip())
                buf = []
            else:
                buf.append(ch)
        if buf:
            parts.append(''.join(buf).strip())

        columns: Dict[str, ColumnDef] = {}
        cinfo = ConstraintInfo()
        for p in parts:
            cm = COLUMN_DEF_RE.match(p)
            if cm:
                col = strip_quotes(cm.group("col"))
                dtype = cm.group("type").strip()
                nullspec = cm.group("null")
                default = cm.group("default")
                columns[col] = ColumnDef(
                    name=col,
                    data_type=dtype,
                    nullable=(False if nullspec and nullspec.upper() == "NOT NULL" else True),
                    default=(default.strip() if default else None),
                )
                continue
            pm = PK_RE.search(p)
            if pm:
                cinfo.primary_key = parse_ident_list(pm.group("cols"))
                continue
            um = UNIQUE_RE.search(p)
            if um:
                c = parse_ident_list(um.group("cols"))
                cinfo.unique_constraints.append(c)
                continue
            fm = FK_RE.search(p)
            if fm:
                cols = parse_ident_list(fm.group("cols"))
                reftab = strip_quotes(fm.group("reftab"))
                if '.' in reftab:
                    rs, rt = [strip_quotes(x) for x in reftab.split('.', 1)]
                else:
                    rs, rt = ("public", reftab)
                refcols = parse_ident_list(fm.group("refcols"))
                cinfo.foreign_keys.append((cols, f'{rs}.{rt}', refcols))
                continue
            # Ignore other constraints for now

        tables[table] = TableDef(schema=schema, name=table, columns=columns, constraints=cinfo)
    return tables


# ---------------------------
# Dependency ordering
# ---------------------------

def topo_sort_tables(tables: Dict[str, TableDef]) -> List[str]:
    # Build directed graph: edge A->B if A depends on B (A has FK to B)
    graph: Dict[str, set[str]] = {t: set() for t in tables.keys()}
    indeg: Dict[str, int] = {t: 0 for t in tables.keys()}
    for tname, tdef in tables.items():
        for (_local_cols, refqual, _refcols) in tdef.constraints.foreign_keys:
            ref_table = refqual.split('.')[-1]
            # Only consider if referenced table exists in our set (same dump)
            if ref_table in tables:
                if ref_table not in graph[tname]:
                    graph[tname].add(ref_table)
    # compute indegrees
    indeg = {t: 0 for t in tables.keys()}
    for _, deps in graph.items():
        for d in deps:
            indeg[d] += 1
    # Kahn
    order: List[str] = []
    S: List[str] = sorted([t for t, d in indeg.items() if d == 0])
    while S:
        n: str = S.pop(0)
        order.append(n)
        for m in list(graph[n]):
            indeg[m] -= 1
            if indeg[m] == 0:
                S.append(m)
        S.sort()
    # If cycle, append remaining nodes alphabetically
    if len(order) < len(tables):
        remaining: List[str] = sorted([t for t in tables.keys() if t not in order])
        order.extend(remaining)
    return order


# ---------------------------
# Excel Reading & Coercion
# ---------------------------

UNSUPPORTED_TYPE_PATTERNS = [
    r"\bbytea\b",
    r"\[[^\]]+\]",  # arrays
]

ENUM_TYPE_RE = re.compile(r"^\"?[A-Za-z_][A-Za-z0-9_]*\"?\.\"?[A-Za-z_][A-Za-z0-9_]*\"?$")  # schema.enumname


def detect_unsupported_type(dtype: str) -> Optional[str]:
    dt = dtype.strip().lower()
    for pat in UNSUPPORTED_TYPE_PATTERNS:
        if re.search(pat, dt):
            return dtype
    # crude enum detection: schema.enumname or unqualified enum
    if not any(kw in dt for kw in ["char", "varchar", "text", "int", "bigint", "smallint", "numeric", "decimal", "double", "real", "bool", "date", "timestamp", "uuid", "jsonb"]):
        # Might be an enum or custom type
        return dtype
    return None


NUMERIC_PREC_SCALE_RE = re.compile(r"(?:numeric|decimal)\s*\(\s*(\d+)\s*,\s*(\d+)\s*\)", re.IGNORECASE)
VARCHAR_LEN_RE = re.compile(r"(?:character varying|varchar)\s*\(\s*(\d+)\s*\)", re.IGNORECASE)


def extract_varchar_len(dtype: str) -> Optional[int]:
    m = VARCHAR_LEN_RE.search(dtype)
    return int(m.group(1)) if m else None

def sql_literal_json_text(s: str) -> str:
    # For JSON text we must NOT double backslashes; just escape single quotes for SQL
    return "'" + s.replace("'", "''") + "'"

def extract_numeric_scale(dtype: str) -> Optional[int]:
    m = NUMERIC_PREC_SCALE_RE.search(dtype)
    return int(m.group(2)) if m else None


def is_double_precision(dtype: str) -> bool:
    return dtype.strip().lower().startswith("double precision") or dtype.strip().lower().startswith("float8")


def is_real(dtype: str) -> bool:
    return dtype.strip().lower().startswith("real") or dtype.strip().lower().startswith("float4")


def is_integerish(dtype: str) -> bool:
    dt = dtype.strip().lower()
    return any(dt.startswith(x) for x in ["smallint", "integer", "bigint", "int2", "int4", "int8", "serial", "bigserial", "smallserial"])


def is_numeric(dtype: str) -> bool:
    return dtype.strip().lower().startswith(("numeric", "decimal"))


def is_boolean(dtype: str) -> bool:
    return dtype.strip().lower().startswith(("bool", "boolean"))


def is_textual(dtype: str) -> bool:
    dt = dtype.strip().lower()
    return ("text" in dt) or ("character varying" in dt) or ("varchar" in dt) or ("char" in dt)

def is_json(dtype: str) -> bool:
    dt = dtype.strip().lower()
    return dt.startswith("json")  # matches 'json' and 'jsonb'

def is_uuid(dtype: str) -> bool:
    return dtype.strip().lower().startswith("uuid")


def is_date(dtype: str) -> bool:
    dt = dtype.strip().lower()
    return dt == "date"


def is_timestamp_tz(dtype: str) -> bool:
    dt = dtype.strip().lower()
    return "timestamp with time zone" in dt or dt.startswith("timestamptz")


def is_timestamp_no_tz(dtype: str) -> bool:
    dt = dtype.strip().lower()
    return "timestamp without time zone" in dt or (dt.startswith("timestamp") and "with time zone" not in dt)

def escape_control_chars_inside_strings(json_like: str) -> str:
    """
    Replace raw newline/tab characters *inside quoted strings* with \\n / \\t.
    Leaves whitespace/newlines outside strings untouched (pretty JSON remains valid).
    """
    out = []
    in_str = False
    escape = False
    for ch in json_like.replace('\r', '\n'):  # normalize CRLF -> LF
        if in_str:
            if escape:
                out.append(ch)
                escape = False
            else:
                if ch == '\\':
                    out.append(ch)
                    escape = True
                elif ch == '"':
                    out.append(ch)
                    in_str = False
                elif ch == '\n':
                    out.append('\\n')
                elif ch == '\t':
                    out.append('\\t')
                else:
                    out.append(ch)
        else:
            if ch == '"':
                out.append(ch)
                in_str = True
            else:
                out.append(ch)
    return ''.join(out)

def coerce_cell(col: ColumnDef, raw: Any, report: List[str], rownum: int) -> Tuple[str, bool]:
    """
    Return (SQL_value_expression, is_default_used)
    SQL_value_expression can be "DEFAULT" or a literal like '...', 42, TIMESTAMPTZ '...'
    """
    dtype = col.data_type
    # Unsupported?
    unsup = detect_unsupported_type(dtype)
    if unsup:
        raise RuntimeError(f"Unsupported column type detected for column {col.name}: {unsup}")

    # Handle NULL
    if raw is None or (isinstance(raw, float) and math.isnan(raw)) or (isinstance(raw, str) and raw.strip() == ""):
        if col.default is not None and not col.nullable:
            # Use DEFAULT keyword
            return ("DEFAULT", True)
        else:
            # No explicit default: follow per-type fallback rules for NOT NULL; NULL otherwise
            if not col.nullable:
                # NOT NULL fallback
                if is_integerish(dtype) or is_numeric(dtype) or is_real(dtype) or is_double_precision(dtype):
                    return ("0" if not (is_real(dtype) or is_double_precision(dtype) or is_numeric(dtype)) else ("0.0" if (is_real(dtype) or is_double_precision(dtype)) else "0"), False)
                elif is_boolean(dtype):
                    return ("FALSE", False)
                elif is_textual(dtype):
                    return (sql_literal_text(""), False)
                elif is_uuid(dtype):
                    # if NOT NULL uuid and missing, generate UUID
                    new_uuid = str(uuid.uuid4())
                    report.append(f"[Row {rownum}] Missing NOT NULL uuid; generated {new_uuid}")
                    return (sql_literal_text(new_uuid), False)
                elif is_date(dtype) or is_timestamp_no_tz(dtype) or is_timestamp_tz(dtype):
                    # No implicit default for dates/timestamps if not provided and no DEFAULT
                    raise ValueError(f"[Row {rownum}] Missing NOT NULL date/timestamp and no DEFAULT specified in schema.")
                else:
                    raise ValueError(f"[Row {rownum}] Missing NOT NULL value for unsupported type: {dtype}")
            else:
                return ("NULL", False)

    # Non-empty: coerce by type
    if is_boolean(dtype):
        b = normalize_boolean(raw)
        if b is None:
            if not col.nullable and col.default is not None:
                return ("DEFAULT", True)
            elif col.nullable:
                return ("NULL", False)
            else:
                raise ValueError(f"[Row {rownum}] Missing NOT NULL boolean and no DEFAULT")
        return (sql_literal_bool(b), False)

    if is_integerish(dtype):
        try:
            # Normalize string numbers with separators
            s = str(raw).strip()
            s = s.replace(" ", "").replace(",", "")
            n = int(float(s)) if not isinstance(raw, (int,)) else int(raw)
        except Exception:
            raise ValueError(f"[Row {rownum}] Invalid integer: {raw}")
        return (sql_literal_int(n), False)

    if is_numeric(dtype):
        scale = extract_numeric_scale(dtype)
        d = parse_decimal_normalize(raw, scale=scale)
        if d is None:
            if col.default is not None and not col.nullable:
                return ("DEFAULT", True)
            return ("NULL", False)
        return (sql_literal_numeric(d), False)

    if is_real(dtype) or is_double_precision(dtype):
        # treat as float; normalize commas
        s = str(raw).strip().replace(" ", "").replace(",", ".")
        try:
            x = float(s)
        except Exception:
            raise ValueError(f"[Row {rownum}] Invalid floating number: {raw}")
        if is_double_precision(dtype):
            return (sql_literal_float(x), False)
        else:
            return (sql_literal_float(x), False)  # same literal form
        
    if is_json(col.data_type):
        # NULL / DEFAULT handling
        if raw is None or (isinstance(raw, float) and math.isnan(raw)) or (isinstance(raw, str) and raw.strip() == ""):
            if col.default is not None and not col.nullable:
                return ("DEFAULT", True)
            return ("NULL" if col.nullable else ("DEFAULT", True)[0], col.default is not None and not col.nullable)

        # Accept dict/list directly; or parse string
        if isinstance(raw, (dict, list)):
            obj = raw
        else:
            s = str(raw)
            # First pass: try strict JSON
            try:
                obj = json.loads(s)
            except json.JSONDecodeError:
                # Repair common Excel artifacts: raw newlines/tabs inside quoted strings
                repaired = escape_control_chars_inside_strings(s)
                try:
                    obj = json.loads(repaired)
                except json.JSONDecodeError as e2:
                    # Give a precise message with a small excerpt around the error position
                    pos = getattr(e2, 'pos', None)
                    snippet = (repaired[max(0, (pos or 0) - 30):(pos or 0) + 30] if pos is not None else repaired[:60])
                    raise ValueError(f"[Row {rownum}] Invalid JSON for column {col.name}: {e2}. Near: {snippet!r}")

        json_text = json.dumps(obj, ensure_ascii=False)
        return (sql_literal_json_text(json_text), False)  
    
    if is_textual(dtype):
        val = str(raw).strip()
        if val == "":
            if not col.nullable:
                if col.default is not None:
                    return ("DEFAULT", True)
                else:
                    return (sql_literal_text(""), False)
            else:
                return ("NULL", False)
        # length check for varchar(n)
        limit = extract_varchar_len(dtype)
        if limit is not None and len(val) > limit:
            raise ValueError(f"[Row {rownum}] Text length {len(val)} exceeds varchar({limit}) for column {col.name}")
        return (sql_literal_text(val), False)

    if is_uuid(dtype):
        s = str(raw).strip()
        if s == "":
            # same as missing above, but we shouldn't reach here
            new_uuid = str(uuid.uuid4())
            report.append(f"[Row {rownum}] Empty uuid; generated {new_uuid}")
            return (sql_literal_text(new_uuid), False)
        # basic validation
        try:
            _ = uuid.UUID(s)
        except Exception:
            raise ValueError(f"[Row {rownum}] Invalid uuid: {raw}")
        return (sql_literal_text(s), False)

    if is_date(dtype):
        d = parse_date_only(raw)
        if d is None:
            if not col.nullable and col.default is not None:
                return ("DEFAULT", True)
            elif col.nullable:
                return ("NULL", False)
            else:
                raise ValueError(f"[Row {rownum}] Missing NOT NULL date and no DEFAULT")
        return (sql_literal_date(d), False)

    if is_timestamp_tz(dtype):
        dt = parse_timestamp_utc(raw, assume_noon_if_date_only=True)
        if dt is None:
            if not col.nullable and col.default is not None:
                return ("DEFAULT", True)
            elif col.nullable:
                return ("NULL", False)
            else:
                raise ValueError(f"[Row {rownum}] Missing NOT NULL timestamptz and no DEFAULT")
        return (sql_literal_timestamp_utc(dt), False)

    if is_timestamp_no_tz(dtype):
        dt = parse_timestamp_utc(raw, assume_noon_if_date_only=True)
        if dt is None:
            if not col.nullable and col.default is not None:
                return ("DEFAULT", True)
            elif col.nullable:
                return ("NULL", False)
            else:
                raise ValueError(f"[Row {rownum}] Missing NOT NULL timestamp and no DEFAULT")
        # Drop tz info but keep UTC instant
        dt_naive = dt.astimezone(timezone.utc).replace(tzinfo=None)
        return (sql_literal_timestamp_naive(dt_naive), False)

    # Fallback: unsupported
    raise RuntimeError(f"Unsupported or unrecognized type for column {col.name}: {dtype}")


# ---------------------------
# Main generation logic
# ---------------------------

def find_excel_file(folder: Path) -> Path:
    cands = sorted(list(folder.glob("*.xlsx"))) + sorted(list(folder.glob("*.xls")))
    if not cands:
        raise FileNotFoundError(f"No Excel file (*.xlsx|*.xls) found in {folder}")
    return cands[0]


def load_schema_file(folder: Path) -> str:
    path = folder / "schema.txt"
    if not path.exists():
        raise FileNotFoundError(f"schema.txt not found in {folder}")
    return path.read_text(encoding="utf-8")


def choose_conflict_target(tdef: TableDef) -> List[str]:
    if tdef.constraints.primary_key:
        return tdef.constraints.primary_key
    if tdef.constraints.unique_constraints:
        return tdef.constraints.unique_constraints[0]
    return []  # no upsert possible


def build_insert_batches(tdef: TableDef, df: pd.DataFrame, report: List[str]) -> Tuple[List[str], int, int]:
    """
    Returns (list_of_insert_statements, success_rows, failed_rows)
    """
    # Ensure df columns match schema columns intersection
    table_cols = list(tdef.columns.keys())

    # Prepare coercion and value assembly per row
    col_objs = [tdef.columns[c] for c in table_cols]
    qualified = q_qualified(tdef.schema, tdef.name)

    conflict_cols = choose_conflict_target(tdef)
    if conflict_cols:
        conflict_target = ", ".join(q_ident(c) for c in conflict_cols)
        non_key_cols = [c for c in table_cols if c not in conflict_cols]
        update_set = ", ".join(f"{q_ident(c)} = EXCLUDED.{q_ident(c)}" for c in non_key_cols)
        on_conflict = f" ON CONFLICT ({conflict_target}) DO UPDATE SET {update_set}"
    else:
        on_conflict = ""

    # Build value tuples
    value_rows: List[str] = []
    success = 0
    failed = 0

    for row_idx, (_, row) in enumerate(df.iterrows()):
        rownum: int = row_idx + 2  # +2 to account for header (row 1) in Excel 1-based
        try:
            vals: List[str] = []
            for col in col_objs:
                raw = row.get(col.name, None)
                valexpr, _ = coerce_cell(col, raw, report, rownum=rownum)
                vals.append(valexpr)
            value_rows.append("(" + ", ".join(vals) + ")")
            success += 1
        except Exception as e:
            report.append(f"[Row {rownum}] ERROR: {e}")
            failed += 1

    # Chunk into batches
    stmts: List[str] = []
    if success > 0:
        col_list = ", ".join(q_ident(c) for c in table_cols)
        for i in range(0, len(value_rows), BATCH_SIZE):
            chunk = value_rows[i:i+BATCH_SIZE]
            stmt = f"INSERT INTO {qualified} ({col_list}) VALUES\n  " + ",\n  ".join(chunk) + on_conflict + ";"
            stmts.append(stmt)

    return stmts, success, failed


def generate_for_table(tdef: TableDef, xl: pd.ExcelFile, report: List[str]) -> Optional[Tuple[str, List[str], int, int]]:
    # Return (sheet_name, statements, success, failed) or None if sheet missing
    sheet_names = xl.sheet_names
    if tdef.name not in sheet_names:
        report.append(f"[WARN] Table '{tdef.name}' in schema has no matching Excel sheet.")
        return None

    df = xl.parse(sheet_name=tdef.name, dtype=str)  # type: ignore
    df: pd.DataFrame  # Explicitly declare type for clarity
    # Trim columns to known schema columns; warn for extras/missing
    known_cols = set(tdef.columns.keys())
    extra_cols = [c for c in df.columns if c not in known_cols]
    if extra_cols:
        report.append(f"[WARN] Sheet '{tdef.name}' has extra columns ignored: {extra_cols}")
    missing_cols = [c for c in known_cols if c not in df.columns]
    if missing_cols:
        # Include missing columns with None values
        for c in missing_cols:
            df[c] = None
        report.append(f"[WARN] Sheet '{tdef.name}' missing columns supplied as NULL/DEFAULT: {missing_cols}")
    # Reorder columns to schema order
    df = df[[c for c in tdef.columns.keys()]]

    # Build statements
    stmts, ok, bad = build_insert_batches(tdef, df, report)
    return (tdef.name, stmts, ok, bad)


def main():
    outdir = OUTPUT_DIR
    outdir.mkdir(parents=True, exist_ok=True)

    report_lines: List[str] = []
    ts = now_utc_stamp_for_filename()

    # Load inputs
    schema_text = load_schema_file(outdir)
    xl_path = find_excel_file(outdir)
    report_lines.append(f"Using Excel file: {xl_path.name}")
    report_lines.append(f"Using schema file: schema.txt")

    # Parse schema
    tables = parse_schema(schema_text)
    if not tables:
        print("No CREATE TABLE statements found in schema.txt", file=sys.stderr)
        sys.exit(1)

    # Unsupported type detection upfront
    for t in tables.values():
        for c in t.columns.values():
            unsup = detect_unsupported_type(c.data_type)
            if unsup:
                print(f"Unsupported type detected in table {t.name}, column {c.name}: {unsup}", file=sys.stderr)
                sys.exit(2)

    # Dependency order
    order = topo_sort_tables(tables)
    report_lines.append("Table order (FK-based, cycles broken alphabetically if any): " + " -> ".join(order))

    # Open Excel
    xl = pd.ExcelFile(xl_path)

    # Warn sheets with no schema
    for s in xl.sheet_names:
        if s not in tables:
            report_lines.append(f"[WARN] Excel sheet '{s}' has no matching table in schema (skipped).")

    # Generate per-table SQL
    results: List[str] = []
    total_ok = 0
    total_bad = 0
    for idx, tname in enumerate(order, start=1):
        tdef = tables[tname]
        res = generate_for_table(tdef, xl, report_lines)
        if res is None:
            continue
        sheet_name, stmts, ok, bad = res
        total_ok += ok
        total_bad += bad

        # Header and file writing
        prefix = f"{idx:02d}"
        fname = f"{prefix}_{sheet_name}_{ts}.sql"
        header = dedent(f"""\
        -- Generated by generate_pg_inserts.py on {ts}
        -- Target Postgres: {PG_VERSION}
        -- Source Excel: {xl_path.name}
        -- Source Schema: schema.txt
        -- Table: {q_qualified(tdef.schema, tdef.name)}
        -- Conflict target: {choose_conflict_target(tdef) or "NONE (no upsert)"}
        -- Rows OK: {ok}, Rows Failed: {bad}
        -- Notes:
        --   * Transaction-wrapped; batched INSERTs with DEFAULT keyword where applicable.
        --   * On conflict: update all non-key columns.
        --   * Timestamps are treated as UTC; date-only inputs -> noon UTC for timestamps; DATE keeps only date.
        --   * Empty strings in text columns -> NULL (if nullable) else '' (empty string).
        --   * Missing NOT NULL date/timestamp without DEFAULT -> row fails (see report.txt).
        """)

        sql_text = "BEGIN;\n\n" + "\n\n".join(stmts) + "\n\nCOMMIT;\n"

        (outdir / fname).write_text(header + "\n" + sql_text, encoding="utf-8")
        results.append(fname)

    # Write report
    report_header = [
        f"Report generated at {ts} (UTC)",
        f"Output directory: {outdir}",
        f"Excel file: {xl_path.name}",
        f"Schema file: schema.txt",
        f"Total rows OK: {total_ok}",
        f"Total rows Failed: {total_bad}",
        "",
        "Messages:",
    ]
    (outdir / "report.txt").write_text("\n".join(report_header + report_lines), encoding="utf-8")

    print("Done.")
    print("Created files:")
    for f in results:
        print(" -", f)
    print(" - report.txt")


if __name__ == "__main__":
    main()
