
# file: app.py
from __future__ import annotations

import io
import re
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl.utils import get_column_letter


st.set_page_config(page_title="Bank Statement Reader", layout="wide")


COLUMN_ALIASES = {
    "trx_date": [
        "tanggal",
        "tgl",
        "date",
        "trx date",
        "transaction date",
        "posting date",
        "tanggal transaksi",
    ],
    "description": [
        "keterangan",
        "deskripsi",
        "uraian",
        "description",
        "transaction description",
        "remark",
        "remarks",
        "narrative",
    ],
    "debit": [
        "debit",
        "debet",
        "db",
    ],
    "credit": [
        "credit",
        "kredit",
        "cr",
    ],
    "balance": [
        "saldo",
        "balance",
        "running balance",
        "saldo akhir",
        "ending balance",
    ],
    "amount": [
        "mutasi",
        "amount",
        "nominal",
        "nilai",
    ],
    "dc": [
        "db/cr",
        "d/c",
        "dk",
        "type",
        "tipe",
        "jenis",
        "posisi",
        "mutasi type",
    ],
    "account_id": [
        "rekening",
        "no rekening",
        "no. rekening",
        "nomor rekening",
        "no rek",
        "norek",
        "rekening no",
        "account",
        "account no",
        "account number",
    ],
    "account_name": [
        "nama rekening",
        "nama account",
        "account name",
        "nama",
        "atas nama",
        "nama nasabah",
    ],
    "opening_balance_explicit": [
        "saldo awal",
        "opening balance",
        "beginning balance",
    ],
}


DEFAULT_MASTER_TEXT = {
    "BCA": """1\t0613419702\tBCA\tBANK BCA CASHLESS BATAM
2\t4301191191\tBCA\tBANK BCA PNP BAKAUHENI
3\t8200827831\tBCA\tBANK BCA CASHLESS SIBOLGA
4\t2950400652\tBCA\tBANK BCA PNP MERAK
5\t2642537777\tBCA\tBANK BCA PNP KETAPANG
6\t1870888828\tBCA\tBANK BCA PNP SURABAYA
7\t7810337154\tBCA\tBANK BCA CASHLESS BALIKPAPAN
8\t8685126334\tBCA\tBANK BCA CASHLESS BATULICIN
9\t7855301644\tBCA\tBANK BCA CASHLESS TERNATE
10\t3141086306\tBCA\tBANK BCA CASHLESS KUPANG
11\t7255999001\tBCA\tBANK BCA PNP KAYANGAN
12\t0561743893\tBCA\tBANK BCA PNP LEMBAR
13\t7065038676\tBCA\tBANK BCA CASHLESS SAPE
14\t8745194440\tBCA\tBANK BCA CASHLESS BAJOE
15\t0411613436\tBCA\tBANK BCA BANGKA
16\t3900925572\tBCA\tBANK BCA SELAYAR
17\t0441598776\tBCA\tBANK BCA AMBON
18\t0223259861\tBCA\tBANK BCA ACEH
19\t0322725645\tBCA\tBANK BCA PADANG
20\t6795136821\tBCA\tBANK BCA LUWUK
21\t6495342828\tBCA\tBANK BCA BAU-BAU
22\t4500553842\tBCA\tBANK BCA PELABUHAN""",
    "Mandiri": "",
    "BNI": "",
}


def normalize_spaces(text: object) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def normalize_column_name(name: object) -> str:
    text = str(name or "").strip().lower()
    text = re.sub(r"[\r\n\t]+", " ", text)
    text = re.sub(r"[_\-]+", " ", text)
    text = re.sub(r"[^\w\s/\.]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def clean_account_id(value: object) -> str:
    raw = normalize_spaces(value)
    digits = re.sub(r"\D", "", raw)
    if len(digits) >= 5:
        return digits
    return raw or "UNKNOWN"


def normalize_account_key(value: object, length: int = 10) -> str:
    raw = clean_account_id(value)
    digits = re.sub(r"\D", "", raw)
    if digits:
        return digits[:length]
    text = normalize_spaces(raw)
    return text[:length] if text else "UNKNOWN"


def display_account_id(value: object) -> str:
    return normalize_account_key(value, length=10)


def format_currency(value: object) -> str:
    if pd.isna(value):
        return ""
    try:
        num = float(value)
    except Exception:
        return str(value)
    return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def parse_amount(value: object) -> Optional[float]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None

    if isinstance(value, (int, float)) and not pd.isna(value):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    negative = False
    if "(" in text and ")" in text:
        negative = True

    text = text.replace("Rp", "").replace("rp", "")
    text = text.replace("IDR", "").replace("idr", "")
    text = text.replace(" ", "")
    text = text.replace("\u00a0", "")
    text = re.sub(r"[^0-9,.\-]", "", text)

    if not text or text in {"-", ".", ","}:
        return None

    if text.startswith("-"):
        negative = True
        text = text[1:]

    if "." in text and "," in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        parts = text.split(",")
        if len(parts[-1]) == 2:
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "." in text:
        parts = text.split(".")
        if len(parts[-1]) == 2 and len(parts) > 1:
            text = text.replace(",", "")
        else:
            text = text.replace(".", "")

    try:
        result = float(text)
    except ValueError:
        return None

    return -result if negative else result


def parse_date_value(value: object, year_hint: Optional[int] = None) -> pd.Timestamp:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return value.normalize()

    text = normalize_spaces(value)
    if not text:
        return pd.NaT

    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y", "%Y-%m-%d"):
        try:
            return pd.Timestamp(datetime.strptime(text, fmt).date())
        except ValueError:
            pass

    simple_dm = re.fullmatch(r"(\d{1,2})/(\d{1,2})", text)
    if simple_dm and year_hint:
        day = int(simple_dm.group(1))
        month = int(simple_dm.group(2))
        try:
            return pd.Timestamp(datetime(year_hint, month, day).date())
        except ValueError:
            return pd.NaT

    parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
    if pd.isna(parsed):
        return pd.NaT
    return pd.Timestamp(parsed.date())


def guess_year_from_text(text: str) -> Optional[int]:
    years = re.findall(r"\b(20\d{2})\b", text)
    if years:
        counts = Counter(int(y) for y in years)
        return counts.most_common(1)[0][0]

    short_years = re.findall(r"\b\d{1,2}/\d{1,2}/(\d{2})\b", text)
    if short_years:
        inferred = [2000 + int(y) for y in short_years]
        counts = Counter(inferred)
        return counts.most_common(1)[0][0]

    return None


def detect_account_id_from_text(text: str, filename: str) -> str:
    patterns = [
        r"(?:NO\.?\s*REKENING|NOMOR\s*REKENING|NO\s*REK(?:ENING)?|ACCOUNT\s*NUMBER)\s*[:\-]?\s*([0-9 \-]{5,})",
        r"(?:REKENING)\s*[:\-]?\s*([0-9 \-]{5,})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return clean_account_id(match.group(1))

    return clean_account_id(Path(filename).stem)


def detect_account_name_from_text(text: str) -> str:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    patterns = [
        r"(?:NAMA\s*NASABAH|NAMA\s*REKENING|ATAS\s*NAMA|ACCOUNT\s*NAME)\s*[:\-]?\s*(.+)$",
        r"(?:NAMA)\s*[:\-]?\s*(.+)$",
    ]
    for line in lines:
        for pattern in patterns:
            match = re.search(pattern, line, flags=re.IGNORECASE)
            if match:
                candidate = normalize_spaces(match.group(1))
                candidate = re.sub(r"[^A-Za-z0-9 .,&/\-]", "", candidate).strip()
                if len(candidate) >= 3:
                    return candidate
    return ""


def detect_opening_balance_from_text(text: str) -> Optional[float]:
    patterns = [
        r"SALDO\s*AWAL\s*[:\-]?\s*([0-9.,]+)",
        r"OPENING\s*BALANCE\s*[:\-]?\s*([0-9.,]+)",
        r"BEGINNING\s*BALANCE\s*[:\-]?\s*([0-9.,]+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return parse_amount(match.group(1))
    return None


def extract_pdf_text(file_bytes: bytes) -> str:
    texts: List[str] = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            texts.append(page.extract_text(x_tolerance=2, y_tolerance=3) or "")
    return "\n".join(texts)


def merge_transaction_lines(lines: List[str]) -> List[str]:
    merged: List[str] = []
    current = ""

    for raw_line in lines:
        line = normalize_spaces(raw_line)
        if not line:
            continue

        if re.match(r"^\d{1,2}/\d{1,2}(?:/\d{2,4})?\b", line):
            if current:
                merged.append(current)
            current = line
        elif current:
            current = f"{current} {line}"

    if current:
        merged.append(current)

    return merged


def detect_dc_marker(body: str, amount_span: Tuple[int, int]) -> Optional[str]:
    start, end = amount_span
    context = body[max(0, start - 20): min(len(body), end + 20)]

    if re.search(r"\b(?:DB|DEBIT|DEBET|D)\b", context, flags=re.IGNORECASE):
        return "DB"
    if re.search(r"\b(?:CR|KREDIT|CREDIT|K)\b", context, flags=re.IGNORECASE):
        return "CR"

    tail = body[end:]
    if re.search(r"^\s*(?:DB|DEBIT|DEBET|D)\b", tail, flags=re.IGNORECASE):
        return "DB"
    if re.search(r"^\s*(?:CR|KREDIT|CREDIT|K)\b", tail, flags=re.IGNORECASE):
        return "CR"

    return None


def parse_pdf_transaction_line(
    line: str,
    year_hint: Optional[int],
    account_id: str,
    account_name: str,
    opening_balance_explicit: Optional[float],
    source_file: str,
    row_order: int,
) -> Optional[Dict[str, object]]:
    line = normalize_spaces(line)
    date_match = re.match(r"^(\d{1,2}/\d{1,2}(?:/\d{2,4})?)\s+(.*)$", line)
    if not date_match:
        return None

    date_text = date_match.group(1)
    body = date_match.group(2).strip()
    if not body:
        return None

    money_pattern = re.compile(r"(?<!\d)(?:\d{1,3}(?:[.,]\d{3})+|\d+)(?:[.,]\d{2})?(?!\d)")
    matches = list(money_pattern.finditer(body))
    if len(matches) < 2:
        return None

    amount_match = matches[-2]
    balance_match = matches[-1]

    amount = parse_amount(amount_match.group(0))
    balance = parse_amount(balance_match.group(0))
    if amount is None and balance is None:
        return None

    dc_marker = detect_dc_marker(body, (amount_match.start(), amount_match.end()))
    description = normalize_spaces(body[: amount_match.start()].strip())
    if not description:
        description = "(tanpa keterangan)"

    debit = 0.0
    credit = 0.0
    if amount is not None:
        if dc_marker == "DB":
            debit = float(abs(amount))
        elif dc_marker == "CR":
            credit = float(abs(amount))

    return {
        "account_id": account_id,
        "account_name": account_name,
        "trx_date": parse_date_value(date_text, year_hint),
        "description": description,
        "amount": float(abs(amount)) if amount is not None else None,
        "debit": debit,
        "credit": credit,
        "balance": balance,
        "opening_balance_explicit": opening_balance_explicit,
        "source_file": source_file,
        "source_sheet": "PDF",
        "row_order": row_order,
        "dc_raw": dc_marker or "",
    }


def infer_missing_debit_credit(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    result = df.sort_values(["trx_date", "row_order"], kind="stable").copy()
    first_opening = result["opening_balance_explicit"].dropna()
    previous_balance = float(first_opening.iloc[0]) if not first_opening.empty else None

    for idx, row in result.iterrows():
        debit = float(row.get("debit", 0) or 0)
        credit = float(row.get("credit", 0) or 0)
        balance = row.get("balance")
        amount = row.get("amount")

        if debit > 0 or credit > 0:
            if pd.notna(balance):
                previous_balance = float(balance)
            continue

        if pd.notna(balance) and previous_balance is not None:
            delta = float(balance) - float(previous_balance)
            guessed = float(abs(amount)) if pd.notna(amount) else abs(delta)
            if delta > 0:
                result.at[idx, "credit"] = guessed
                result.at[idx, "debit"] = 0.0
            elif delta < 0:
                result.at[idx, "debit"] = guessed
                result.at[idx, "credit"] = 0.0
            previous_balance = float(balance)
        elif pd.notna(amount):
            marker = str(row.get("dc_raw", "")).upper()
            if marker == "CR":
                result.at[idx, "credit"] = float(abs(amount))
                result.at[idx, "debit"] = 0.0
            elif marker == "DB":
                result.at[idx, "debit"] = float(abs(amount))
                result.at[idx, "credit"] = 0.0
            if pd.notna(balance):
                previous_balance = float(balance)
        elif pd.notna(balance):
            previous_balance = float(balance)

    return result


def read_csv_with_fallbacks(file_bytes: bytes) -> pd.DataFrame:
    encodings = ["utf-8", "utf-8-sig", "cp1252", "latin1"]
    last_error: Optional[Exception] = None

    for encoding in encodings:
        try:
            return pd.read_csv(io.BytesIO(file_bytes), encoding=encoding)
        except Exception as exc:
            last_error = exc

    raise ValueError(f"Gagal membaca CSV: {last_error}")


def best_matching_column(columns: List[str], aliases: List[str]) -> Optional[str]:
    normalized_cols = {col: normalize_column_name(col) for col in columns}

    for alias in aliases:
        alias_norm = normalize_column_name(alias)
        for original, normalized in normalized_cols.items():
            if normalized == alias_norm:
                return original

    for alias in aliases:
        alias_norm = normalize_column_name(alias)
        for original, normalized in normalized_cols.items():
            if alias_norm in normalized or normalized in alias_norm:
                return original

    return None


def map_columns(df: pd.DataFrame) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    columns = list(df.columns)
    for canonical, aliases in COLUMN_ALIASES.items():
        matched = best_matching_column(columns, aliases)
        if matched:
            mapping[matched] = canonical
    return mapping


def standardize_dc(value: object) -> str:
    text = normalize_spaces(value).upper()
    if text in {"DB", "DEBIT", "DEBET", "D"}:
        return "DB"
    if text in {"CR", "CREDIT", "KREDIT", "K"}:
        return "CR"
    return ""


def first_non_empty(series: pd.Series) -> str:
    for value in series:
        text = normalize_spaces(value)
        if text:
            return text
    return ""


def create_manifest_row(
    account_id: str,
    account_name: str,
    source_file: str,
    source_sheet: str,
    parse_status: str,
    parse_note: str,
    transaction_count: int,
) -> Dict[str, object]:
    account_key = normalize_account_key(account_id)
    return {
        "account_id": account_key,
        "account_name": normalize_spaces(account_name),
        "source_file": source_file,
        "source_sheet": source_sheet,
        "parse_status": parse_status,
        "parse_note": parse_note,
        "transaction_count": int(transaction_count),
    }


def build_manifest_from_transactions(parsed_df: pd.DataFrame, filename: str, sheet_name: str) -> pd.DataFrame:
    if parsed_df.empty:
        return pd.DataFrame(
            columns=[
                "account_id",
                "account_name",
                "source_file",
                "source_sheet",
                "parse_status",
                "parse_note",
                "transaction_count",
            ]
        )

    rows: List[Dict[str, object]] = []
    for account_id, group in parsed_df.groupby("account_id", dropna=False, sort=True):
        rows.append(
            create_manifest_row(
                account_id=str(account_id),
                account_name=first_non_empty(group["account_name"]) if "account_name" in group.columns else "",
                source_file=filename,
                source_sheet=sheet_name,
                parse_status="OK",
                parse_note="Transaksi berhasil dibaca",
                transaction_count=len(group),
            )
        )
    return pd.DataFrame(rows)


def extract_account_hints_from_dataframe(raw_df: pd.DataFrame, filename: str, sheet_name: str) -> List[Dict[str, str]]:
    df = raw_df.copy().dropna(axis=0, how="all").dropna(axis=1, how="all")
    if df.empty:
        return [{"account_id": f"UNKNOWN_{Path(filename).stem}_{sheet_name}", "account_name": ""}]

    mapping = map_columns(df)
    df = df.rename(columns=mapping)

    if "account_id" in df.columns:
        temp = df.copy()
        temp["account_id"] = temp["account_id"].fillna("").astype(str).apply(normalize_spaces)
        if "account_name" in temp.columns:
            temp["account_name"] = temp["account_name"].fillna("").astype(str).apply(normalize_spaces)
        else:
            temp["account_name"] = ""

        rows: List[Dict[str, str]] = []
        seen = set()
        for _, row in temp.iterrows():
            account_key = normalize_account_key(row["account_id"])
            if not account_key or account_key in seen or account_key == "UNKNOWN":
                continue
            seen.add(account_key)
            rows.append({"account_id": account_key, "account_name": normalize_spaces(row["account_name"])})
        if rows:
            return rows

    for value in df.astype(str).fillna("").values.flatten().tolist():
        match = re.search(r"(?<!\d)(\d{8,20})(?!\d)", normalize_spaces(value))
        if match:
            return [{"account_id": normalize_account_key(match.group(1)), "account_name": ""}]

    return [{"account_id": f"UNKNOWN_{Path(filename).stem}_{sheet_name}", "account_name": ""}]


def convert_spreadsheet_to_transactions(
    raw_df: pd.DataFrame,
    filename: str,
    sheet_name: str,
) -> Tuple[pd.DataFrame, List[str]]:
    notes: List[str] = []
    df = raw_df.copy().dropna(axis=0, how="all").dropna(axis=1, how="all")
    if df.empty:
        return pd.DataFrame(), [f"{filename} [{sheet_name}]: sheet kosong"]

    mapping = map_columns(df)
    df = df.rename(columns=mapping)

    if "trx_date" not in df.columns:
        return pd.DataFrame(), [f"{filename} [{sheet_name}]: kolom tanggal tidak ditemukan"]

    optional_defaults = {
        "description": "",
        "account_id": Path(filename).stem,
        "account_name": "",
        "opening_balance_explicit": None,
        "debit": None,
        "credit": None,
        "balance": None,
        "amount": None,
        "dc": "",
    }
    for col, default in optional_defaults.items():
        if col not in df.columns:
            df[col] = default

    df["trx_date"] = df["trx_date"].apply(parse_date_value)
    df["description"] = df["description"].fillna("").astype(str).apply(normalize_spaces)
    df["account_id"] = df["account_id"].apply(clean_account_id)
    df["account_name"] = df["account_name"].fillna("").astype(str).apply(normalize_spaces)
    df["opening_balance_explicit"] = df["opening_balance_explicit"].apply(parse_amount)
    df["debit"] = df["debit"].apply(parse_amount)
    df["credit"] = df["credit"].apply(parse_amount)
    df["balance"] = df["balance"].apply(parse_amount)
    df["amount"] = df["amount"].apply(parse_amount)
    df["dc"] = df["dc"].apply(standardize_dc)

    if df["amount"].notna().any():
        for idx, row in df.iterrows():
            amount = row["amount"]
            if pd.isna(amount):
                continue

            debit = row["debit"] if pd.notna(row["debit"]) else 0.0
            credit = row["credit"] if pd.notna(row["credit"]) else 0.0

            if debit == 0 and credit == 0:
                if row["dc"] == "DB" or amount < 0:
                    df.at[idx, "debit"] = float(abs(amount))
                    df.at[idx, "credit"] = 0.0
                elif row["dc"] == "CR" or amount > 0:
                    df.at[idx, "credit"] = float(abs(amount))
                    df.at[idx, "debit"] = 0.0

    df["debit"] = df["debit"].fillna(0.0).astype(float)
    df["credit"] = df["credit"].fillna(0.0).astype(float)
    df["source_file"] = filename
    df["source_sheet"] = sheet_name
    df["row_order"] = range(len(df))
    df["dc_raw"] = df["dc"]

    required_cols = [
        "account_id",
        "account_name",
        "trx_date",
        "description",
        "amount",
        "debit",
        "credit",
        "balance",
        "opening_balance_explicit",
        "source_file",
        "source_sheet",
        "row_order",
        "dc_raw",
    ]
    df = df[required_cols]
    df = df[df["trx_date"].notna()].copy()

    if df.empty:
        return pd.DataFrame(), [f"{filename} [{sheet_name}]: tidak ada baris transaksi valid"]

    notes.append(
        f"{filename} [{sheet_name}]: sheet terbaca | rekening unik={df['account_id'].nunique()} | transaksi={len(df)}"
    )
    return df, notes


def parse_bca_pdf(file_bytes: bytes, filename: str) -> Tuple[pd.DataFrame, pd.DataFrame, List[str]]:
    notes: List[str] = []

    try:
        text = extract_pdf_text(file_bytes)
    except Exception as exc:
        manifest_df = pd.DataFrame(
            [
                create_manifest_row(
                    account_id=Path(filename).stem,
                    account_name="",
                    source_file=filename,
                    source_sheet="PDF",
                    parse_status="ERROR",
                    parse_note=f"Gagal baca PDF: {exc}",
                    transaction_count=0,
                )
            ]
        )
        return pd.DataFrame(), manifest_df, [f"{filename}: error baca PDF - {exc}"]

    if not text.strip():
        manifest_df = pd.DataFrame(
            [
                create_manifest_row(
                    account_id=Path(filename).stem,
                    account_name="",
                    source_file=filename,
                    source_sheet="PDF",
                    parse_status="NO_TEXT",
                    parse_note="PDF tidak mengandung teks yang bisa diekstrak",
                    transaction_count=0,
                )
            ]
        )
        return pd.DataFrame(), manifest_df, [f"{filename}: PDF tidak mengandung teks yang bisa diekstrak"]

    year_hint = guess_year_from_text(text)
    account_id = detect_account_id_from_text(text, filename)
    account_name = detect_account_name_from_text(text)
    opening_balance_explicit = detect_opening_balance_from_text(text)

    rows: List[Dict[str, object]] = []
    merged_lines = merge_transaction_lines([line for line in text.splitlines() if line.strip()])
    for row_order, line in enumerate(merged_lines):
        parsed = parse_pdf_transaction_line(
            line=line,
            year_hint=year_hint,
            account_id=account_id,
            account_name=account_name,
            opening_balance_explicit=opening_balance_explicit,
            source_file=filename,
            row_order=row_order,
        )
        if parsed is not None:
            rows.append(parsed)

    if not rows:
        manifest_df = pd.DataFrame(
            [
                create_manifest_row(
                    account_id=account_id,
                    account_name=account_name,
                    source_file=filename,
                    source_sheet="PDF",
                    parse_status="NO_TRANSACTION",
                    parse_note="No rekening terdeteksi, tetapi transaksi tidak ditemukan",
                    transaction_count=0,
                )
            ]
        )
        return pd.DataFrame(), manifest_df, [f"{filename}: no rekening terdeteksi, transaksi PDF tidak ditemukan"]

    df = infer_missing_debit_credit(pd.DataFrame(rows))
    manifest_df = build_manifest_from_transactions(df, filename, "PDF")
    notes.append(f"{filename}: PDF terbaca | rekening={normalize_account_key(account_id)} | transaksi={len(df)}")
    return df, manifest_df, notes


def parse_tabular_file(file_bytes: bytes, filename: str, ext: str) -> Tuple[pd.DataFrame, pd.DataFrame, List[str]]:
    notes: List[str] = []
    results: List[pd.DataFrame] = []
    manifests: List[pd.DataFrame] = []

    if ext == ".csv":
        raw_df = read_csv_with_fallbacks(file_bytes)
        parsed_df, df_notes = convert_spreadsheet_to_transactions(raw_df, filename, "CSV")
        notes.extend(df_notes)

        if not parsed_df.empty:
            results.append(parsed_df)
            manifests.append(build_manifest_from_transactions(parsed_df, filename, "CSV"))
        else:
            hints = extract_account_hints_from_dataframe(raw_df, filename, "CSV")
            manifests.append(
                pd.DataFrame(
                    [
                        create_manifest_row(
                            account_id=hint["account_id"],
                            account_name=hint["account_name"],
                            source_file=filename,
                            source_sheet="CSV",
                            parse_status="NO_TRANSACTION",
                            parse_note="File terbaca tetapi tidak ada transaksi valid",
                            transaction_count=0,
                        )
                        for hint in hints
                    ]
                )
            )
    else:
        workbook = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)

        if not workbook:
            manifests.append(
                pd.DataFrame(
                    [
                        create_manifest_row(
                            account_id=Path(filename).stem,
                            account_name="",
                            source_file=filename,
                            source_sheet="WORKBOOK",
                            parse_status="EMPTY_WORKBOOK",
                            parse_note="Workbook kosong",
                            transaction_count=0,
                        )
                    ]
                )
            )

        for sheet_name, raw_df in workbook.items():
            parsed_df, df_notes = convert_spreadsheet_to_transactions(raw_df, filename, str(sheet_name))
            notes.extend(df_notes)

            if not parsed_df.empty:
                results.append(parsed_df)
                manifests.append(build_manifest_from_transactions(parsed_df, filename, str(sheet_name)))
            else:
                cleaned = raw_df.dropna(axis=0, how="all").dropna(axis=1, how="all")
                hints = extract_account_hints_from_dataframe(raw_df, filename, str(sheet_name))
                status = "EMPTY_SHEET" if cleaned.empty else "NO_TRANSACTION"
                note = "Sheet kosong" if cleaned.empty else "Sheet terbaca tetapi tidak ada transaksi valid"

                manifests.append(
                    pd.DataFrame(
                        [
                            create_manifest_row(
                                account_id=hint["account_id"],
                                account_name=hint["account_name"],
                                source_file=filename,
                                source_sheet=str(sheet_name),
                                parse_status=status,
                                parse_note=note,
                                transaction_count=0,
                            )
                            for hint in hints
                        ]
                    )
                )

    tx_df = pd.concat(results, ignore_index=True) if results else pd.DataFrame()
    manifest_df = pd.concat(manifests, ignore_index=True) if manifests else pd.DataFrame(
        columns=[
            "account_id",
            "account_name",
            "source_file",
            "source_sheet",
            "parse_status",
            "parse_note",
            "transaction_count",
        ]
    )
    return tx_df, manifest_df, notes


def finalize_transactions(df: pd.DataFrame, deduplicate: bool) -> pd.DataFrame:
    if df.empty:
        return df

    result = df.copy()
    result["account_id"] = result["account_id"].fillna("").astype(str).apply(normalize_account_key)
    result["account_name"] = result["account_name"].fillna("").astype(str).apply(normalize_spaces)
    result["description"] = result["description"].fillna("").astype(str).apply(normalize_spaces)
    result["debit"] = pd.to_numeric(result["debit"], errors="coerce").fillna(0.0)
    result["credit"] = pd.to_numeric(result["credit"], errors="coerce").fillna(0.0)
    result["balance"] = pd.to_numeric(result["balance"], errors="coerce")
    result["amount"] = pd.to_numeric(result["amount"], errors="coerce")
    result["opening_balance_explicit"] = pd.to_numeric(result["opening_balance_explicit"], errors="coerce")

    result = result.sort_values(
        by=["account_id", "trx_date", "source_file", "source_sheet", "row_order"],
        kind="stable",
    ).reset_index(drop=True)

    if deduplicate:
        result = result.drop_duplicates(
            subset=["account_id", "trx_date", "description", "debit", "credit", "balance"],
            keep="first",
        ).reset_index(drop=True)

    return result


def parse_master_accounts(selected_bank: str) -> pd.DataFrame:
    master_text = DEFAULT_MASTER_TEXT.get(selected_bank, "")
    rows: List[Dict[str, object]] = []

    for line in str(master_text or "").splitlines():
        clean_line = normalize_spaces(line)
        if not clean_line:
            continue

        parts = re.split(r"\t+|\s{2,}", line.strip())
        if len(parts) >= 4 and parts[0].isdigit():
            order_str, account_id, bank_name, account_desc = parts[0], parts[1], parts[2], " ".join(parts[3:])
        else:
            tokens = clean_line.split(" ", 3)
            if len(tokens) < 4:
                continue
            order_str, account_id, bank_name, account_desc = tokens[0], tokens[1], tokens[2], tokens[3]

        rows.append(
            {
                "master_order": int(re.sub(r"\D", "", order_str) or len(rows) + 1),
                "account_id": normalize_account_key(account_id),
                "bank": normalize_spaces(bank_name or selected_bank) or selected_bank,
                "master_name": normalize_spaces(account_desc),
            }
        )

    if not rows:
        return pd.DataFrame(columns=["master_order", "account_id", "bank", "master_name"])

    master_df = pd.DataFrame(rows).drop_duplicates(subset=["account_id"], keep="first")
    master_df = master_df.sort_values("master_order", kind="stable").reset_index(drop=True)
    return master_df


def resolve_account_name(account_key: str, summary_name: str, manifest_name: str, master_df: pd.DataFrame) -> str:
    if master_df is not None and not master_df.empty:
        matched = master_df.loc[master_df["account_id"] == account_key, "master_name"]
        if not matched.empty and normalize_spaces(matched.iloc[0]):
            return normalize_spaces(matched.iloc[0])

    if normalize_spaces(summary_name):
        return normalize_spaces(summary_name)

    return normalize_spaces(manifest_name)


def sort_transactions(group: pd.DataFrame) -> pd.DataFrame:
    sort_cols = [col for col in ["trx_date", "source_file", "source_sheet", "row_order"] if col in group.columns]
    return group.sort_values(sort_cols, kind="stable").reset_index(drop=True)


def derive_first_balance_opening(first_row: pd.Series) -> Optional[float]:
    balance = first_row.get("balance")
    if pd.isna(balance):
        return None
    debit = float(first_row.get("debit", 0) or 0)
    credit = float(first_row.get("credit", 0) or 0)
    return float(balance) + debit - credit


def derive_opening_balance(group: pd.DataFrame) -> float:
    explicit = group["opening_balance_explicit"].dropna()
    if not explicit.empty:
        return float(explicit.iloc[0])

    sorted_group = sort_transactions(group)
    balance_rows = sorted_group[sorted_group["balance"].notna()]
    if not balance_rows.empty:
        opening = derive_first_balance_opening(balance_rows.iloc[0])
        if opening is not None:
            return opening

    return 0.0


def derive_closing_balance(group: pd.DataFrame, opening_balance: float) -> float:
    sorted_group = sort_transactions(group)
    balance_rows = sorted_group[sorted_group["balance"].notna()]
    if not balance_rows.empty:
        return float(balance_rows.iloc[-1]["balance"])

    total_debit = float(sorted_group["debit"].sum())
    total_credit = float(sorted_group["credit"].sum())
    return opening_balance + total_credit - total_debit


def build_summary_base(df: pd.DataFrame) -> pd.DataFrame:
    records: List[Dict[str, object]] = []

    if not df.empty:
        for account_key, group in df.groupby("account_id", dropna=False, sort=True):
            ordered = sort_transactions(group)
            records.append(
                {
                    "Rekening": normalize_account_key(account_key),
                    "Nama Rekening": first_non_empty(ordered["account_name"]) if "account_name" in ordered.columns else "",
                    "Saldo Awal": derive_opening_balance(ordered),
                    "Debit": float(ordered["debit"].sum()),
                    "Kredit": float(ordered["credit"].sum()),
                    "Saldo Akhir": derive_closing_balance(ordered, derive_opening_balance(ordered)),
                    "Jumlah Transaksi": len(ordered),
                }
            )

    return pd.DataFrame(
        records,
        columns=[
            "Rekening",
            "Nama Rekening",
            "Saldo Awal",
            "Debit",
            "Kredit",
            "Saldo Akhir",
            "Jumlah Transaksi",
        ],
    )


def finalize_summary(
    summary_df: pd.DataFrame,
    manifest_df: pd.DataFrame,
    master_df: pd.DataFrame,
) -> pd.DataFrame:
    summary_df = summary_df.copy()
    if summary_df.empty:
        summary_df = pd.DataFrame(
            columns=[
                "Rekening",
                "Nama Rekening",
                "Saldo Awal",
                "Debit",
                "Kredit",
                "Saldo Akhir",
                "Jumlah Transaksi",
            ]
        )

    summary_df["Rekening"] = summary_df["Rekening"].astype(str).apply(normalize_account_key)

    manifest_accounts = pd.DataFrame(columns=["account_id", "account_name"])
    if manifest_df is not None and not manifest_df.empty:
        manifest_accounts = (
            manifest_df[["account_id", "account_name"]]
            .fillna("")
            .copy()
        )
        manifest_accounts["account_id"] = manifest_accounts["account_id"].astype(str).apply(normalize_account_key)
        manifest_accounts["account_name"] = manifest_accounts["account_name"].astype(str).apply(normalize_spaces)
        manifest_accounts = manifest_accounts.drop_duplicates(subset=["account_id"], keep="first").reset_index(drop=True)

    keys = set(summary_df["Rekening"].astype(str).tolist())
    if not manifest_accounts.empty:
        keys.update(manifest_accounts["account_id"].astype(str).tolist())
    if master_df is not None and not master_df.empty:
        keys.update(master_df["account_id"].astype(str).tolist())

    result_rows: List[Dict[str, object]] = []
    summary_map = summary_df.set_index("Rekening").to_dict("index") if not summary_df.empty else {}
    manifest_name_map = manifest_accounts.set_index("account_id")["account_name"].to_dict() if not manifest_accounts.empty else {}
    master_order_map = dict(zip(master_df["account_id"], master_df["master_order"])) if master_df is not None and not master_df.empty else {}

    for account_key in keys:
        row = summary_map.get(account_key, {})
        result_rows.append(
            {
                "Rekening": display_account_id(account_key),
                "Nama Rekening": resolve_account_name(
                    account_key=account_key,
                    summary_name=row.get("Nama Rekening", ""),
                    manifest_name=manifest_name_map.get(account_key, ""),
                    master_df=master_df,
                ),
                "Saldo Awal": float(row.get("Saldo Awal", 0.0) or 0.0),
                "Debit": float(row.get("Debit", 0.0) or 0.0),
                "Kredit": float(row.get("Kredit", 0.0) or 0.0),
                "Saldo Akhir": float(row.get("Saldo Akhir", 0.0) or 0.0),
                "Jumlah Transaksi": int(row.get("Jumlah Transaksi", 0) or 0),
                "_master_order": master_order_map.get(account_key, 999999),
                "_account_key": account_key,
            }
        )

    result = pd.DataFrame(result_rows)
    if result.empty:
        return pd.DataFrame(
            columns=[
                "Rekening",
                "Nama Rekening",
                "Saldo Awal",
                "Debit",
                "Kredit",
                "Saldo Akhir",
                "Jumlah Transaksi",
            ]
        )

    result = result.sort_values(["_master_order", "_account_key"], kind="stable").reset_index(drop=True)
    result = result.drop(columns=["_master_order", "_account_key"])
    return result


def build_summary(
    transactions: pd.DataFrame,
    manifest_df: pd.DataFrame,
    master_df: pd.DataFrame,
) -> pd.DataFrame:
    summary_base = build_summary_base(transactions)
    return finalize_summary(summary_base, manifest_df, master_df)


def derive_daily_account_opening(account_history: pd.DataFrame, day_rows: pd.DataFrame) -> float:
    ordered_history = sort_transactions(account_history)
    ordered_day = sort_transactions(day_rows)

    explicit_before_day = ordered_history[
        ordered_history["opening_balance_explicit"].notna() & (ordered_history["trx_date"] <= ordered_day.iloc[0]["trx_date"])
    ]["opening_balance_explicit"]
    explicit_value = float(explicit_before_day.iloc[0]) if not explicit_before_day.empty else None

    first_row = ordered_day.iloc[0]
    first_opening = derive_first_balance_opening(first_row)
    if first_opening is not None:
        return first_opening

    previous_rows = ordered_history[
        (ordered_history["trx_date"] < first_row["trx_date"])
        | (
            (ordered_history["trx_date"] == first_row["trx_date"])
            & (
                (ordered_history["source_file"] < first_row["source_file"])
                | (
                    (ordered_history["source_file"] == first_row["source_file"])
                    & (
                        (ordered_history["source_sheet"] < first_row["source_sheet"])
                        | (
                            (ordered_history["source_sheet"] == first_row["source_sheet"])
                            & (ordered_history["row_order"] < first_row["row_order"])
                        )
                    )
                )
            )
        )
    ]

    previous_balances = previous_rows[previous_rows["balance"].notna()]
    if not previous_balances.empty:
        return float(previous_balances.iloc[-1]["balance"])

    if explicit_value is not None:
        return explicit_value

    return 0.0


def derive_daily_account_closing(day_rows: pd.DataFrame, day_opening: float) -> float:
    ordered_day = sort_transactions(day_rows)
    balance_rows = ordered_day[ordered_day["balance"].notna()]
    if not balance_rows.empty:
        return float(balance_rows.iloc[-1]["balance"])
    return day_opening + float(ordered_day["credit"].sum()) - float(ordered_day["debit"].sum())


def build_daily_summary_map(
    transactions: pd.DataFrame,
    manifest_df: pd.DataFrame,
    master_df: pd.DataFrame,
) -> Dict[str, pd.DataFrame]:
    if transactions.empty or "trx_date" not in transactions.columns:
        return {}

    valid = transactions[transactions["trx_date"].notna()].copy()
    if valid.empty:
        return {}

    daily_records: Dict[str, List[Dict[str, object]]] = {}

    for account_key, account_history in valid.groupby("account_id", dropna=False, sort=True):
        ordered_history = sort_transactions(account_history)
        for trx_date, day_rows in ordered_history.groupby(ordered_history["trx_date"].dt.date, sort=True):
            ordered_day = sort_transactions(day_rows)
            day_opening = derive_daily_account_opening(ordered_history, ordered_day)
            day_closing = derive_daily_account_closing(ordered_day, day_opening)
            date_key = str(trx_date)

            if date_key not in daily_records:
                daily_records[date_key] = []

            daily_records[date_key].append(
                {
                    "Rekening": normalize_account_key(account_key),
                    "Nama Rekening": first_non_empty(ordered_day["account_name"]) if "account_name" in ordered_day.columns else "",
                    "Saldo Awal": day_opening,
                    "Debit": float(ordered_day["debit"].sum()),
                    "Kredit": float(ordered_day["credit"].sum()),
                    "Saldo Akhir": day_closing,
                    "Jumlah Transaksi": len(ordered_day),
                }
            )

    result: Dict[str, pd.DataFrame] = {}
    for date_key, rows in daily_records.items():
        day_summary = pd.DataFrame(rows)
        day_summary = finalize_summary(day_summary, manifest_df, master_df)
        result[date_key] = day_summary[["Rekening", "Saldo Awal", "Debit", "Kredit", "Saldo Akhir"]].copy()

    return result


def make_display_copy(df: pd.DataFrame, money_columns: List[str]) -> pd.DataFrame:
    display_df = df.copy()
    for col in money_columns:
        if col in display_df.columns:
            display_df[col] = display_df[col].apply(format_currency)
    return display_df


def sanitize_sheet_name(name: str, used_names: set[str]) -> str:
    clean = re.sub(r"[:\\/?*\[\]]", "_", str(name)).strip()
    clean = clean[:31] or "Sheet"

    if clean not in used_names:
        used_names.add(clean)
        return clean

    base = clean[:28] or "Sht"
    counter = 1
    while True:
        candidate = f"{base}_{counter}"[:31]
        if candidate not in used_names:
            used_names.add(candidate)
            return candidate
        counter += 1


def autosize_worksheet(ws) -> None:
    for column_cells in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))
        ws.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 40)


def build_excel_export(
    summary_df: pd.DataFrame,
    detail_df: pd.DataFrame,
    manifest_df: pd.DataFrame,
    daily_summary_map: Dict[str, pd.DataFrame],
) -> bytes:
    output = io.BytesIO()
    used_sheet_names: set[str] = set()

    detail_export = detail_df.copy()
    if not detail_export.empty and "Tanggal" in detail_export.columns:
        detail_export["Tanggal"] = pd.to_datetime(detail_export["Tanggal"], errors="coerce")
        detail_export["Tanggal"] = detail_export["Tanggal"].dt.strftime("%Y-%m-%d")

    status_export = manifest_df.copy()
    if status_export.empty:
        status_export = pd.DataFrame(
            columns=[
                "account_id",
                "account_name",
                "source_file",
                "source_sheet",
                "parse_status",
                "parse_note",
                "transaction_count",
            ]
        )

    status_export = status_export.rename(
        columns={
            "account_id": "Rekening",
            "account_name": "Nama Rekening",
            "source_file": "File",
            "source_sheet": "Sheet",
            "parse_status": "Status",
            "parse_note": "Catatan",
            "transaction_count": "Jumlah Transaksi",
        }
    )
    if "Rekening" in status_export.columns:
        status_export["Rekening"] = status_export["Rekening"].astype(str).apply(display_account_id)
    status_export = status_export.sort_values(["File", "Sheet", "Rekening"], kind="stable").reset_index(drop=True)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name=sanitize_sheet_name("Rekap", used_names=used_sheet_names), index=False)
        detail_export.to_excel(writer, sheet_name=sanitize_sheet_name("Semua_Transaksi", used_names=used_sheet_names), index=False)
        status_export.to_excel(writer, sheet_name=sanitize_sheet_name("Status_File", used_names=used_sheet_names), index=False)

        for date_key, day_df in daily_summary_map.items():
            day_df.to_excel(writer, sheet_name=sanitize_sheet_name(date_key, used_names=used_sheet_names), index=False)

        workbook = writer.book
        money_columns = {"Saldo Awal", "Debit", "Kredit", "Saldo Akhir", "Saldo"}

        for ws in workbook.worksheets:
            headers = [cell.value for cell in ws[1]]
            for col_idx, header in enumerate(headers, start=1):
                if header in money_columns:
                    for row_idx in range(2, ws.max_row + 1):
                        ws.cell(row=row_idx, column=col_idx).number_format = "#,##0.00"
            autosize_worksheet(ws)

    output.seek(0)
    return output.getvalue()


def main() -> None:
    st.title("Bank Statement Reader")
    st.caption("Upload banyak file sekaligus, gabungkan banyak rekening, dan export rekap per rekening / per tanggal.")

    with st.sidebar:
        st.subheader("Opsi")
        selected_bank = st.selectbox("Bank", ["BCA", "Mandiri", "BNI"], index=0)
        deduplicate = st.checkbox("Hapus duplikat transaksi identik", value=True)
        st.markdown(
            """
            **Format file**
            - PDF
            - CSV
            - XLSX / XLS

            **Catatan**
            - parser PDF saat ini paling cocok untuk layout BCA berbasis teks
            - master rekening default disimpan di code, tidak ditampilkan di UI
            """
        )

    master_df = parse_master_accounts(selected_bank)

    uploaded_files = st.file_uploader(
        "Pilih file mutasi / rekening koran",
        type=["pdf", "csv", "xlsx", "xls"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        st.info("Upload file dulu untuk mulai proses.")
        return

    parsed_dfs: List[pd.DataFrame] = []
    manifest_dfs: List[pd.DataFrame] = []
    parser_notes: List[str] = []
    parser_errors: List[str] = []

    progress = st.progress(0, text="Memproses file...")
    total_files = len(uploaded_files)

    for i, uploaded_file in enumerate(uploaded_files, start=1):
        filename = uploaded_file.name
        ext = Path(filename).suffix.lower()
        file_bytes = uploaded_file.getvalue()

        try:
            if ext == ".pdf":
                if selected_bank != "BCA":
                    df_file = pd.DataFrame()
                    manifest_file = pd.DataFrame(
                        [
                            create_manifest_row(
                                account_id=Path(filename).stem,
                                account_name="",
                                source_file=filename,
                                source_sheet="PDF",
                                parse_status="UNSUPPORTED_PDF",
                                parse_note=f"Parser PDF khusus {selected_bank} belum tersedia",
                                transaction_count=0,
                            )
                        ]
                    )
                    notes = [f"{filename}: parser PDF khusus {selected_bank} belum tersedia"]
                else:
                    df_file, manifest_file, notes = parse_bca_pdf(file_bytes, filename)
            elif ext in {".csv", ".xlsx", ".xls"}:
                df_file, manifest_file, notes = parse_tabular_file(file_bytes, filename, ext)
            else:
                df_file = pd.DataFrame()
                manifest_file = pd.DataFrame(
                    [
                        create_manifest_row(
                            account_id=Path(filename).stem,
                            account_name="",
                            source_file=filename,
                            source_sheet="FILE",
                            parse_status="UNSUPPORTED",
                            parse_note="Format file tidak didukung",
                            transaction_count=0,
                        )
                    ]
                )
                notes = [f"{filename}: format file tidak didukung"]

            parser_notes.extend(notes)

            if not manifest_file.empty:
                manifest_dfs.append(manifest_file)

            if not df_file.empty:
                parsed_dfs.append(df_file)
            else:
                parser_errors.append(f"{filename}: tidak ada transaksi valid, tetapi file tetap dicatat")
        except Exception as exc:
            parser_errors.append(f"{filename}: error - {exc}")
            manifest_dfs.append(
                pd.DataFrame(
                    [
                        create_manifest_row(
                            account_id=Path(filename).stem,
                            account_name="",
                            source_file=filename,
                            source_sheet="FILE",
                            parse_status="ERROR",
                            parse_note=str(exc),
                            transaction_count=0,
                        )
                    ]
                )
            )

        progress.progress(i / total_files, text=f"Memproses file {i}/{total_files}: {filename}")

    progress.empty()

    if not manifest_dfs:
        st.error("Tidak ada file yang berhasil dicatat.")
        return

    manifest_df = pd.concat(manifest_dfs, ignore_index=True)
    manifest_df["account_id"] = manifest_df["account_id"].astype(str).apply(normalize_account_key)

    transactions = pd.concat(parsed_dfs, ignore_index=True) if parsed_dfs else pd.DataFrame()
    if not transactions.empty:
        transactions = finalize_transactions(transactions, deduplicate=deduplicate)

    summary_df = build_summary(transactions, manifest_df=manifest_df, master_df=master_df)
    daily_summary_map = build_daily_summary_map(transactions, manifest_df=manifest_df, master_df=master_df)

    if not transactions.empty:
        detail_df = transactions[
            [
                "account_id",
                "account_name",
                "trx_date",
                "description",
                "debit",
                "credit",
                "balance",
                "source_file",
                "source_sheet",
                "row_order",
            ]
        ].rename(
            columns={
                "account_id": "Rekening",
                "account_name": "Nama Rekening",
                "trx_date": "Tanggal",
                "description": "Keterangan",
                "debit": "Debit",
                "credit": "Kredit",
                "balance": "Saldo",
                "source_file": "File",
                "source_sheet": "Sheet",
                "row_order": "Row",
            }
        )
        detail_df["Rekening"] = detail_df["Rekening"].astype(str).apply(display_account_id)
        master_name_map = dict(zip(master_df["account_id"], master_df["master_name"])) if not master_df.empty else {}
        detail_df["Nama Rekening"] = detail_df.apply(
            lambda row: normalize_spaces(master_name_map.get(normalize_account_key(row["Rekening"]), "")) or normalize_spaces(row["Nama Rekening"]),
            axis=1,
        )
    else:
        detail_df = pd.DataFrame(
            columns=["Rekening", "Nama Rekening", "Tanggal", "Keterangan", "Debit", "Kredit", "Saldo", "File", "Sheet", "Row"]
        )

    st.success(
        f"Upload: {len(uploaded_files)} file | "
        f"Tercatat: {manifest_df['source_file'].nunique()} file | "
        f"Transaksi valid: {len(transactions) if not transactions.empty else 0} baris"
    )

    st.subheader("Rekap per Rekening")
    st.dataframe(
        make_display_copy(summary_df, ["Saldo Awal", "Debit", "Kredit", "Saldo Akhir"]),
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("Status File / Sheet")
    status_preview = manifest_df.rename(
        columns={
            "account_id": "Rekening",
            "account_name": "Nama Rekening",
            "source_file": "File",
            "source_sheet": "Sheet",
            "parse_status": "Status",
            "parse_note": "Catatan",
            "transaction_count": "Jumlah Transaksi",
        }
    )
    status_preview["Rekening"] = status_preview["Rekening"].astype(str).apply(display_account_id)
    st.dataframe(status_preview, use_container_width=True, hide_index=True)

    st.subheader("Detail Transaksi")
    st.dataframe(
        make_display_copy(detail_df, ["Debit", "Kredit", "Saldo"]),
        use_container_width=True,
        hide_index=True,
    )

    excel_bytes = build_excel_export(
        summary_df=summary_df,
        detail_df=detail_df,
        manifest_df=manifest_df,
        daily_summary_map=daily_summary_map,
    )

    st.download_button(
        label="Download Excel",
        data=excel_bytes,
        file_name=f"rekap_{selected_bank.lower()}_split_per_tanggal.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.download_button(
        label="Download Rekap CSV",
        data=summary_df.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"rekap_{selected_bank.lower()}.csv",
        mime="text/csv",
    )

    st.download_button(
        label="Download Detail CSV",
        data=detail_df.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"detail_{selected_bank.lower()}.csv",
        mime="text/csv",
    )

    with st.expander("Detail log parser"):
        for note in parser_notes:
            st.write(f"- {note}")
        for err in parser_errors:
            st.write(f"- {err}")

    st.caption(
        "Saldo akhir per sheet tanggal diambil dari kolom Saldo pada baris terakhir di tanggal tersebut. "
        "Jika ada beberapa baris pada tanggal yang sama, sistem memilih row terbesar setelah data diurutkan."
    )


if __name__ == "__main__":
    main()
