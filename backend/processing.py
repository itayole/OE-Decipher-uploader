import io
import zipfile
from dataclasses import dataclass, field

from openpyxl import load_workbook

CODING_SHEET_CANDIDATES = ["קידוד", "coding"]
SAMPLE_SIZE = 20


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


def _find_coding_sheet(wb):
    for name in wb.sheetnames:
        if name.strip() in CODING_SHEET_CANDIDATES:
            return wb[name]
    return wb[wb.sheetnames[0]]


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


def process_workbook(file_bytes: bytes) -> dict:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = _find_coding_sheet(wb)
    blocks = detect_blocks(ws)

    results = []
    dat_files = {}
    for block in blocks:
        if block.code_count == 0:
            continue
        columns, rows = build_block_rows(ws, block)
        filename = f"{block.name}_coded.dat"
        dat_files[filename] = rows_to_dat_bytes(columns, rows)
        results.append(
            {
                "question_name": block.name,
                "code_count": block.code_count,
                "row_count": len(rows),
                "filename": filename,
                "columns": columns,
                "preview_rows": rows[:20],
            }
        )

    return {"blocks": results, "dat_files": dat_files}


def zip_dat_files(dat_files: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename, data in dat_files.items():
            zf.writestr(filename, data)
    return buf.getvalue()
