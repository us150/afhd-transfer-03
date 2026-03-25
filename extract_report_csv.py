#!/usr/bin/env python3
"""
Extract Russian accounting line-item codes and current-period amounts from PDF.

Output CSV format:
Код статьи;Сумма
"""

from __future__ import annotations

import argparse
import csv
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ALLOWED_CODES = {
    # Balance sheet (0710001)
    1100,
    1105,
    1110,
    1120,
    1130,
    1140,
    1150,
    1160,
    1170,
    1180,
    1190,
    1200,
    1210,
    1215,
    1220,
    1230,
    1240,
    1250,
    1260,
    1300,
    1310,
    1320,
    1340,
    1350,
    1360,
    1370,
    1400,
    1410,
    1420,
    1430,
    1450,
    1500,
    1510,
    1520,
    1530,
    1540,
    1550,
    1600,
    1700,
    # Profit and loss (0710002)
    2100,
    2110,
    2120,
    2200,
    2210,
    2220,
    2300,
    2310,
    2320,
    2330,
    2340,
    2350,
    2400,
    2410,
    2411,
    2412,
    2420,
    2460,
}

DIGIT_MAP = str.maketrans(
    {
        "О": "0",
        "о": "0",
        "O": "0",
        "З": "3",
        "з": "3",
        "I": "1",
        "l": "1",
        "|": "1",
        "!": "1",
        "]": "1",
        "[": "1",
        "а": "0",
    }
)

MONTHS_RU = {
    "января": "01",
    "февраля": "02",
    "марта": "03",
    "апреля": "04",
    "мая": "05",
    "июня": "06",
    "июля": "07",
    "августа": "08",
    "сентября": "09",
    "октября": "10",
    "ноября": "11",
    "декабря": "12",
}


def ensure_tool(name: str) -> None:
    if shutil.which(name) is None:
        raise RuntimeError(f"Required tool is missing: {name}")


def run(cmd: list[str]) -> None:
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def normalize_token(token: str) -> str:
    return token.translate(DIGIT_MAP)


def only_digits(s: str) -> str:
    return "".join(ch for ch in s if ch.isdigit())


def extract_code_with_index(tokens: list[str]) -> tuple[int, int] | None:
    for idx, tok in enumerate(tokens):
        digits = only_digits(normalize_token(tok))
        if len(digits) != 4:
            continue
        code = int(digits)
        if code in ALLOWED_CODES:
            return code, idx
    return None


def column_count_for_code(code: int) -> int:
    return 3 if 1000 <= code < 2000 else 2


def clean_amount_tokens(tokens_after_code: list[str]) -> list[str]:
    raw_tokens: list[str] = []
    for raw in tokens_after_code:
        tok = normalize_token(raw).strip().strip(",.;:")
        if not tok:
            continue
        tok = tok.replace("−", "-").replace("—", "-").replace("–", "-")

        # OCR artifact: "1-" / "-1" often means a dash placeholder.
        if re.fullmatch(r"\d-", tok) or re.fullmatch(r"-\d", tok):
            tok = "-"

        if tok == "_":
            tok = "-"

        # Keep tokens required for amount parsing (digits, dashes, parentheses).
        if tok in {"(", ")", "-"}:
            raw_tokens.append(tok)
            continue
        if any(ch.isdigit() for ch in tok) or "(" in tok or ")" in tok:
            raw_tokens.append(tok)

    # Merge spaced parentheses forms: ( 123 ) -> (123), ( - ) -> (-)
    out: list[str] = []
    i = 0
    while i < len(raw_tokens):
        tok = raw_tokens[i]
        if tok != "(":
            out.append(tok)
            i += 1
            continue

        j = i + 1
        inner: list[str] = []
        while j < len(raw_tokens) and raw_tokens[j] != ")":
            inner.append(raw_tokens[j])
            j += 1

        if j >= len(raw_tokens):
            # Unbalanced opening parenthesis.
            i += 1
            continue

        inner_join = "".join(inner).replace(" ", "").strip()
        if inner_join in {"", "-", "_"}:
            out.append("(-)")
        elif any(ch.isdigit() for ch in inner_join):
            out.append(f"({inner_join})")
        # else: ignore non-numeric bracket groups.

        i = j + 1

    return out


def fill_derived_values(code_amounts: dict[int, int]) -> dict[int, int]:
    out = dict(code_amounts)

    # Profit (loss) from sales: 2200 = 2100 + 2210 + 2220
    # (commercial and admin expenses are stored with negative sign).
    if 2200 not in out and 2100 in out and 2210 in out and 2220 in out:
        out[2200] = out[2100] + out[2210] + out[2220]

    return out


def parse_column_candidates(tokens: list[str], start: int) -> list[tuple[int | None, int, int]]:
    tok = tokens[start]

    # Dash placeholder means numeric zero in this domain.
    if tok in {"-", "(-)", "()"}:
        return [(0, start + 1, 1)]

    # Parenthesized amount means negative value (with or without spaces originally).
    if tok.startswith("(") and tok.endswith(")"):
        inner = tok[1:-1].strip()
        if inner in {"", "-", "_"}:
            return [(0, start + 1, 1)]
        d_inner = only_digits(inner)
        if not d_inner:
            return []
        return [(-int(d_inner), start + 1, 1)]

    d0 = only_digits(tok)
    if not d0:
        return []

    negative = tok.startswith("-") or tok.startswith("(") or tok.endswith(")")

    out: list[tuple[int | None, int, int]] = []

    base = int(d0)
    if negative:
        base = -base
    out.append((base, start + 1, 1))

    # OCR may merge first thousand group: e.g. "1943 449" -> 1 943 449.
    can_group = len(d0) <= 3 or len(d0) == 4
    if not can_group:
        return out

    max_extra = 0
    j = start + 1
    while j < len(tokens) and max_extra < 2:
        if tokens[j] in {"-", "(-)", "()"}:
            break
        nd = only_digits(tokens[j])
        if len(nd) != 3:
            break
        max_extra += 1
        j += 1

    for extra in range(1, max_extra + 1):
        parts = [d0]
        neg = negative
        ok = True
        for k in range(1, extra + 1):
            t = tokens[start + k]
            if t in {"-", "(-)", "()"}:
                ok = False
                break
            nd = only_digits(t)
            if len(nd) != 3:
                ok = False
                break
            parts.append(nd)
            if ")" in t:
                neg = True
        if not ok:
            continue
        v = int("".join(parts))
        if neg:
            v = -v
        out.append((v, start + 1 + extra, 1 + extra))

    return out


def parse_first_amount(tokens_after_code: list[str], code: int) -> int | None:
    tokens = clean_amount_tokens(tokens_after_code)
    if not tokens:
        return None

    explicit_zero_marker_first = tokens[0] in {"-", "(-)", "()"}

    ncols = column_count_for_code(code)
    parses: list[tuple[list[int | None], list[int], int]] = []

    def dfs(col: int, idx: int, values: list[int | None], groups: list[int]) -> None:
        if col == ncols:
            parses.append((values[:], groups[:], idx))
            return

        if idx >= len(tokens):
            values.append(None)
            groups.append(0)
            dfs(col + 1, idx, values, groups)
            values.pop()
            groups.pop()
            return

        for value, next_idx, group_count in parse_column_candidates(tokens, idx):
            values.append(value)
            groups.append(group_count)
            dfs(col + 1, next_idx, values, groups)
            values.pop()
            groups.pop()

        # Allow missing non-first columns when OCR skipped a value.
        if col > 0:
            values.append(None)
            groups.append(0)
            dfs(col + 1, idx, values, groups)
            values.pop()
            groups.pop()

    dfs(0, 0, [], [])

    best_first: int | None = None
    best_score: tuple[int, int, int, int, int, int] | None = None

    for values, groups, end_idx in parses:
        first = values[0]
        if first is None:
            continue

        present = [v for v in values if v is not None]
        digit_lens = [len(str(abs(v))) for v in present]
        max_digits = max(digit_lens) if digit_lens else 0
        too_long = sum(1 for d in digit_lens if d > 8)
        missing = sum(1 for v in values if v is None)

        used_groups = [g for g, v in zip(groups, values) if v is not None]
        spread = (max(used_groups) - min(used_groups)) if used_groups else 0

        next_chunk_len = len(only_digits(tokens[1])) if len(tokens) > 1 else 0
        small_first_penalty = 0
        if (
            not explicit_zero_marker_first
            and code not in {1220, 2320}
            and abs(first) < 100
            and next_chunk_len == 3
        ):
            small_first_penalty = 1

        leftover = len(tokens) - end_idx
        score = (leftover, too_long, small_first_penalty, missing, max_digits, spread)
        if best_score is None or score < best_score:
            best_score = score
            best_first = first

    return best_first


def extract_reporting_date(text: str) -> str | None:
    m = re.search(r"\bна\s+(\d{1,2})\s+([а-яё]+)\s+(\d{4})\s*г", text.lower())
    if not m:
        return None
    day, month_word, year = m.group(1), m.group(2), m.group(3)
    month = MONTHS_RU.get(month_word)
    if not month:
        return None
    return f"{year}-{month}-{day.zfill(2)}"


def extract_code_amounts(text: str) -> dict[int, int]:
    out: dict[int, int] = {}

    token_lines: list[list[str]] = []
    code_hits: list[tuple[int, int] | None] = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        tokens = line.split() if line else []
        token_lines.append(tokens)
        code_hits.append(extract_code_with_index(tokens) if tokens else None)

    for i, tokens in enumerate(token_lines):
        if not tokens:
            continue

        found = code_hits[i]
        if not found:
            continue

        code, idx = found
        value = parse_first_amount(tokens[idx + 1 :], code)

        # Some scans split table rows into two lines: code on one line, amounts on the next.
        if value is None:
            probes = 0
            j = i + 1
            while j < len(token_lines) and probes < 2:
                next_tokens = token_lines[j]
                next_found = code_hits[j]
                j += 1

                if not next_tokens:
                    continue
                if next_found is not None:
                    break

                probes += 1
                value = parse_first_amount(next_tokens, code)
                if value is not None:
                    break

        if value is None:
            continue

        out.setdefault(code, value)

    return out


def extract_text_from_pdf(pdf_path: Path) -> str:
    ensure_tool("pdftotext")
    result = subprocess.run(
        ["pdftotext", "-layout", str(pdf_path), "-"],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="ignore",
    )
    return result.stdout


def ocr_pdf(pdf_path: Path, dpi: int, lang: str, psm: int) -> str:
    ensure_tool("pdftoppm")
    ensure_tool("tesseract")

    with tempfile.TemporaryDirectory(prefix="pdf_ocr_") as td:
        tmp = Path(td)
        prefix = tmp / "page"

        run(["pdftoppm", "-r", str(dpi), "-png", str(pdf_path), str(prefix)])
        pngs = sorted(tmp.glob("page-*.png"), key=lambda p: int(re.search(r"(\d+)", p.stem).group(1)))

        texts: list[str] = []
        for png in pngs:
            stem = png.with_suffix("")
            run(["tesseract", str(png), str(stem), "-l", lang, "--psm", str(psm)])
            texts.append(stem.with_suffix(".txt").read_text(encoding="utf-8", errors="ignore"))

        return "\n".join(texts)


def write_csv(path: Path, code_amounts: dict[int, int]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Код статьи", "Сумма"])
        for code in sorted(code_amounts):
            w.writerow([code, code_amounts[code]])



def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Batch extract accounting codes from all PDF files located near this script"
    )
    p.add_argument("--dpi", type=int, default=300, help="OCR rasterization DPI (default: 300)")
    p.add_argument("--lang", default="rus", help="Tesseract language (default: rus)")
    p.add_argument("--psm", type=int, default=4, help="Tesseract page segmentation mode (default: 4)")
    p.add_argument(
        "--suffix",
        default=".csv",
        help="Output suffix for each PDF filename (default: .csv)",
    )
    p.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing output CSV files",
    )
    return p


def process_pdf(pdf_path: Path, output_csv: Path, dpi: int, lang: str, psm: int) -> tuple[bool, str]:
    try:
        # 1) OCR extraction (primary path for scanned/partially noisy PDFs).
        ocr_text = ocr_pdf(pdf_path, dpi=dpi, lang=lang, psm=psm)
        ocr_data = extract_code_amounts(ocr_text)

        # 2) Text-layer extraction (supplemental, may recover missed lines).
        text_layer = ""
        text_data: dict[int, int] = {}
        try:
            text_layer = extract_text_from_pdf(pdf_path)
            text_data = extract_code_amounts(text_layer)
        except RuntimeError:
            # pdftotext may be absent; OCR output is enough to continue.
            pass

        code_amounts = dict(text_data)
        code_amounts.update(ocr_data)  # Prefer OCR values when both exist.
        code_amounts = fill_derived_values(code_amounts)

        if not code_amounts:
            return False, "no code/value pairs extracted"

        write_csv(output_csv, code_amounts)

        report_date = extract_reporting_date(text_layer) or extract_reporting_date(ocr_text)
        details = f"rows={len(code_amounts)}"
        if report_date:
            details += f", date={report_date}"
        return True, details
    except subprocess.CalledProcessError as e:
        return False, f"external tool failed: {e}"
    except RuntimeError as e:
        return False, str(e)


def main() -> int:
    args = build_arg_parser().parse_args()

    script_dir = Path(__file__).resolve().parent
    pdf_files = sorted(script_dir.glob("*.pdf"))

    if not pdf_files:
        print(f"No PDF files found in: {script_dir}", file=sys.stderr)
        return 2

    ok_count = 0
    fail_count = 0
    skip_count = 0

    for pdf_path in pdf_files:
        output_csv = script_dir / f"{pdf_path.stem}{args.suffix}"

        if output_csv.exists() and not args.overwrite:
            skip_count += 1
            print(f"SKIP {pdf_path.name} -> {output_csv.name} (already exists)")
            continue

        success, info = process_pdf(
            pdf_path=pdf_path,
            output_csv=output_csv,
            dpi=args.dpi,
            lang=args.lang,
            psm=args.psm,
        )

        if success:
            ok_count += 1
            print(f"OK   {pdf_path.name} -> {output_csv.name} ({info})")
        else:
            fail_count += 1
            print(f"FAIL {pdf_path.name} ({info})", file=sys.stderr)

    print(f"Done: ok={ok_count}, failed={fail_count}, skipped={skip_count}")
    return 0 if fail_count == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
