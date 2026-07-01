import io
import re
import zipfile
from dataclasses import dataclass, field

from openpyxl import load_workbook

CODING_SHEET_CANDIDATES = ["קידוד", "coding"]
CATEGORY_SHEET_CANDIDATES = ["קטגוריות", "categories"]
SAMPLE_SIZE = 20

BLOCK_TYPE_REGULAR = "regular"
BLOCK_TYPE_AB = "ab"
BLOCK_TYPE_CLOSED_OTHERS = "closed_others"


def normalize(value):
    """Normalize a cell value for comparison: blanks -> None, numeric-looking
    values -> a canonical string (so 2, "2", 2.0 all compare equal)."""
    if value is None:
        return None
    s = str(value).strip()
    if s == "":
        return None
    try:
        f = float(s)
        return str(int(f)) if f.is_integer() else str(f)
    except ValueError:
        return s


def _is_numeric_or_blank(value) -> bool:
    if value is None:
        return True
    s = str(value).strip()
    if s == "":
        return True
    try:
        float(s)
        return True
    except ValueError:
        return False


@dataclass
class QuestionBlock:
    name: str
    raw_col_index: int
    code_col_indices: list = field(default_factory=list)

    @property
    def code_count(self) -> int:
        return len(self.code_col_indices)


def _find_sheet(wb, candidates, fallback_index=0):
    for name in wb.sheetnames:
        if name.strip() in candidates:
            return wb[name]
    if fallback_index is not None and fallback_index < len(wb.sheetnames):
        return wb[wb.sheetnames[fallback_index]]
    return None


def _column_is_raw_text(col_values) -> bool:
    """A column starts a new question block if any of its first
    non-blank sampled values is not purely numeric (i.e. free text)."""
    sampled = 0
    for value in col_values:
        if value is None:
            continue
        s = str(value).strip()
        if s == "":
            continue
        sampled += 1
        if not _is_numeric_or_blank(s):
            return True
        if sampled >= SAMPLE_SIZE:
            break
    return False


def detect_blocks(ws) -> list[QuestionBlock]:
    max_row = ws.max_row
    max_col = ws.max_column
    header = [ws.cell(row=1, column=c).value for c in range(1, max_col + 1)]

    blocks: list[QuestionBlock] = []
    current: QuestionBlock | None = None

    for col in range(3, max_col + 1):
        col_values = (ws.cell(row=r, column=col).value for r in range(2, max_row + 1))
        is_raw = _column_is_raw_text(col_values) or current is None
        if is_raw:
            current = QuestionBlock(name=str(header[col - 1]), raw_col_index=col)
            blocks.append(current)
        else:
            current.code_col_indices.append(col)

    return blocks


def load_categories(wb) -> dict:
    """שאלה -> [{code, label}], read from the קטגוריות sheet."""
    ws = _find_sheet(wb, CATEGORY_SHEET_CANDIDATES, fallback_index=None)
    if ws is None:
        return {}
    categories: dict[str, list] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 3:
            continue
        question, code, label = row[0], row[1], row[2]
        if question is None or code is None:
            continue
        question = str(question).strip()
        categories.setdefault(question, []).append({"code": normalize(code), "label": label})
    return categories


_AB_SUFFIX_RE = re.compile(r"^(.*?)([ab])$", re.IGNORECASE)


def suggest_ab_pairs(blocks: list[QuestionBlock]) -> dict:
    """Returns {block_name: {"base": str, "role": "A"|"B", "paired_with": str}}
    for blocks whose name ends in a/b and whose sibling (same base, opposite
    letter) also exists among the detected blocks."""
    by_name = {b.name: b for b in blocks}
    suggestions = {}
    for block in blocks:
        m = _AB_SUFFIX_RE.match(block.name)
        if not m:
            continue
        base, letter = m.group(1), m.group(2).lower()
        sibling_letter = "b" if letter == "a" else "a"
        sibling_name = base + sibling_letter
        sibling = by_name.get(sibling_name) or by_name.get(sibling_name.upper())
        if sibling is None:
            for candidate_name in by_name:
                cm = _AB_SUFFIX_RE.match(candidate_name)
                if cm and cm.group(1) == base and cm.group(2).lower() == sibling_letter:
                    sibling = by_name[candidate_name]
                    sibling_name = candidate_name
                    break
        if sibling is not None:
            suggestions[block.name] = {
                "base": base,
                "role": "A" if letter == "a" else "B",
                "paired_with": sibling_name,
            }
    return suggestions


def count_answered(ws, block: QuestionBlock) -> int:
    """Number of respondents with at least one non-blank code in this block."""
    max_row = ws.max_row
    count = 0
    for r in range(2, max_row + 1):
        if any(normalize(ws.cell(row=r, column=c).value) is not None for c in block.code_col_indices):
            count += 1
    return count


def count_answered_rows(rows: list[list]) -> int:
    """Same as count_answered, for already-built [record, uuid, *codes] rows."""
    return sum(1 for row in rows if any(normalize(v) is not None for v in row[2:]))


def build_block_rows(ws, block: QuestionBlock) -> tuple[list[str], list[list]]:
    max_row = ws.max_row
    columns = ["record", "uuid"] + [f"code{i+1}" for i in range(block.code_count)]
    rows = []
    for r in range(2, max_row + 1):
        record = ws.cell(row=r, column=1).value
        uuid = ws.cell(row=r, column=2).value
        codes = [ws.cell(row=r, column=c).value for c in block.code_col_indices]
        rows.append([record, uuid] + codes)
    return columns, rows


def rows_to_dat_bytes(columns: list[str], rows: list[list]) -> bytes:
    lines = ["\t".join(columns)]
    for row in rows:
        cells = ["" if v is None else str(v) for v in row]
        lines.append("\t".join(cells))
    return ("\r\n".join(lines) + "\r\n").encode("utf-8")


def clean_ab_pair(ws, block_a: QuestionBlock, block_b: QuestionBlock, clean_code: str):
    """Applies the unaided-awareness (TOM / others) cleaning rules to a single
    A/B pair, row by row. Returns (rows_a, rows_b, log_rows)."""
    max_row = ws.max_row
    clean_code = normalize(clean_code)

    def is_blank(v):
        return normalize(v) is None

    def is_clean(v):
        return normalize(v) == clean_code

    def is_real(v):
        return not is_blank(v) and not is_clean(v)

    columns_a = ["record", "uuid"] + [f"code{i+1}" for i in range(block_a.code_count)]
    columns_b = ["record", "uuid"] + [f"code{i+1}" for i in range(block_b.code_count)]
    rows_a, rows_b, log_rows = [], [], []

    for r in range(2, max_row + 1):
        record = ws.cell(row=r, column=1).value
        uid = ws.cell(row=r, column=2).value
        a = [ws.cell(row=r, column=c).value for c in block_a.code_col_indices]
        b = [ws.cell(row=r, column=c).value for c in block_b.code_col_indices]

        def log(rule, block_label, col_idx, old, new):
            log_rows.append(
                [record, uid, rule, block_label, f"code{col_idx + 1}", old, new if new is not None else ""]
            )

        # Rules 1+2: clear cleaned-answer occurrences from B when either the
        # cleaned answer appears in both A and B, or A holds a real answer.
        a_has_real = any(is_real(v) for v in a)
        a_has_clean = any(is_clean(v) for v in a)
        b_has_clean = any(is_clean(v) for v in b)
        if (a_has_clean and b_has_clean) or a_has_real:
            for i, v in enumerate(b):
                if is_clean(v):
                    log("1/2", "B", i, v, None)
                    b[i] = None

        # Rule 3: A's first col is cleaned but B's first col has a real
        # answer -> move that answer into A's first col.
        if block_a.code_count and block_b.code_count and is_clean(a[0]) and is_real(b[0]):
            log("3", "A", 0, a[0], b[0])
            log("3", "B", 0, b[0], None)
            a[0] = b[0]
            b[0] = None

        # Rule 4: within each block, a real first-col answer means any
        # cleaned answer in the other cols of that same block is erased.
        for block_label, arr in (("A", a), ("B", b)):
            if arr and is_real(arr[0]):
                for i in range(1, len(arr)):
                    if is_clean(arr[i]):
                        log("4", block_label, i, arr[i], None)
                        arr[i] = None

        rows_a.append([record, uid] + a)
        rows_b.append([record, uid] + b)

    return (columns_a, rows_a), (columns_b, rows_b), log_rows


def log_rows_to_bytes(log_rows: list[list]) -> bytes:
    header = ["record", "uuid", "rule", "block", "column", "old_value", "new_value"]
    lines = ["\t".join(header)]
    for row in log_rows:
        cells = ["" if v is None else str(v) for v in row]
        lines.append("\t".join(cells))
    return ("\r\n".join(lines) + "\r\n").encode("utf-8")


def detect_and_describe(file_bytes: bytes) -> dict:
    """Step 1: parse the workbook and return blocks with suggested types and
    their category dictionaries, for the mapping-confirmation screen."""
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = _find_sheet(wb, CODING_SHEET_CANDIDATES)
    blocks = [b for b in detect_blocks(ws) if b.code_count > 0]
    categories = load_categories(wb)
    ab_suggestions = suggest_ab_pairs(blocks)

    results = []
    for block in blocks:
        suggestion = ab_suggestions.get(block.name)
        results.append(
            {
                "name": block.name,
                "code_count": block.code_count,
                "answered_count": count_answered(ws, block),
                "suggested_type": BLOCK_TYPE_AB if suggestion else BLOCK_TYPE_REGULAR,
                "paired_with": suggestion["paired_with"] if suggestion else None,
                "role": suggestion["role"] if suggestion else None,
                "categories": categories.get(block.name, []),
            }
        )
    return {"blocks": results}


def generate_outputs(file_bytes: bytes, mapping: list[dict]) -> dict:
    """Step 2: apply the confirmed type mapping and produce the .dat / log
    files. `mapping` is a list of {name, type, cleaned_code?}."""
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = _find_sheet(wb, CODING_SHEET_CANDIDATES)
    blocks = {b.name: b for b in detect_blocks(ws) if b.code_count > 0}
    mapping_by_name = {m["name"]: m for m in mapping if m["name"] in blocks}

    results = []
    dat_files = {}
    warnings = []
    handled = set()

    ab_entries = [m for m in mapping_by_name.values() if m.get("type") == BLOCK_TYPE_AB]
    ab_suggestions = suggest_ab_pairs(list(blocks.values()))

    for entry in ab_entries:
        name = entry["name"]
        if name in handled:
            continue
        suggestion = ab_suggestions.get(name)
        if not suggestion:
            warnings.append(f"{name}: לא נמצא זוג AB תואם (שם עם סיומת a/b) — יטופל כשאלה רגילה")
            continue
        paired_name = suggestion["paired_with"]
        paired_entry = mapping_by_name.get(paired_name)
        if not paired_entry or paired_entry.get("type") != BLOCK_TYPE_AB:
            warnings.append(f"{name}: השאלה המזווגת {paired_name} לא סומנה כ-AB — יטופל כשאלה רגילה")
            continue

        clean_code = entry.get("cleaned_code") or paired_entry.get("cleaned_code")
        if not clean_code:
            warnings.append(f"{name}/{paired_name}: לא נבחרה 'תשובה לניקוי' — יטופל כשאלה רגילה")
            continue

        block_a = blocks[name] if suggestion["role"] == "A" else blocks[paired_name]
        block_b = blocks[paired_name] if suggestion["role"] == "A" else blocks[name]
        (cols_a, rows_a), (cols_b, rows_b), log_rows = clean_ab_pair(ws, block_a, block_b, clean_code)

        base = suggestion["base"]
        fname_a = f"{block_a.name}_coded.dat"
        fname_b = f"{block_b.name}_coded.dat"
        fname_log = f"{base}_cleaning_log.txt"
        dat_files[fname_a] = rows_to_dat_bytes(cols_a, rows_a)
        dat_files[fname_b] = rows_to_dat_bytes(cols_b, rows_b)
        dat_files[fname_log] = log_rows_to_bytes(log_rows)

        results.append(
            {
                "question_name": block_a.name,
                "type": BLOCK_TYPE_AB,
                "role": "A",
                "code_count": block_a.code_count,
                "row_count": len(rows_a),
                "answered_count": count_answered_rows(rows_a),
                "filename": fname_a,
                "columns": cols_a,
                "preview_rows": rows_a[:20],
            }
        )
        results.append(
            {
                "question_name": block_b.name,
                "type": BLOCK_TYPE_AB,
                "role": "B",
                "code_count": block_b.code_count,
                "row_count": len(rows_b),
                "answered_count": count_answered_rows(rows_b),
                "filename": fname_b,
                "columns": cols_b,
                "preview_rows": rows_b[:20],
            }
        )
        results.append(
            {
                "question_name": base,
                "type": "log",
                "code_count": None,
                "row_count": len(log_rows),
                "answered_count": None,
                "filename": fname_log,
                "columns": ["record", "uuid", "rule", "block", "column", "old_value", "new_value"],
                "preview_rows": log_rows[:20],
            }
        )
        handled.add(name)
        handled.add(paired_name)

    for name, block in blocks.items():
        if name in handled:
            continue
        columns, rows = build_block_rows(ws, block)
        filename = f"{block.name}_coded.dat"
        dat_files[filename] = rows_to_dat_bytes(columns, rows)
        entry = mapping_by_name.get(name, {})
        results.append(
            {
                "question_name": block.name,
                "type": entry.get("type", BLOCK_TYPE_REGULAR),
                "role": None,
                "code_count": block.code_count,
                "row_count": len(rows),
                "answered_count": count_answered_rows(rows),
                "filename": filename,
                "columns": columns,
                "preview_rows": rows[:20],
            }
        )

    return {"blocks": results, "dat_files": dat_files, "warnings": warnings}


def zip_dat_files(dat_files: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename, data in dat_files.items():
            zf.writestr(filename, data)
    return buf.getvalue()
