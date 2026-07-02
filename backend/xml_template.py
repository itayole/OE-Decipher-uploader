import re
from pathlib import Path

DEFAULT_TEMPLATE_PATH = Path(__file__).parent / "templates" / "default_template.txt"

DELIMITER_RE = re.compile(r"^[-—]{3,}\s*$")
PLACEHOLDER_RE = re.compile(r'(\w+)\s*=\s*File\(\s*"\1\.dat"\s*,\s*"record"\s*\)')
TITLE_RE = re.compile(r"<title>.*?</title>", re.DOTALL)
ROW_RE = re.compile(r'^(?P<indent>[ \t]*)<row label="r(?P<num>\d+)">(?P<val>.*?)</row>\s*$')
CODE_SLOT_RE = re.compile(r"(?i)^code\d+$")
CHECKBOX_BLOCK_RE = re.compile(r"<checkbox\b.*?</checkbox>", re.DOTALL | re.IGNORECASE)
VIRTUAL_BLOCK_RE = re.compile(r"<virtual>.*?</virtual>", re.DOTALL)


def read_default_template() -> str:
    return DEFAULT_TEMPLATE_PATH.read_text(encoding="utf-8")


def write_default_template(content: str) -> None:
    DEFAULT_TEMPLATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    DEFAULT_TEMPLATE_PATH.write_text(content, encoding="utf-8")


DEFAULT_DELIMITER_LINE = "-" * 72


def extract_template_block(template_text: str) -> str:
    """The template block lives between the first two delimiter lines
    (a line of 3+ dashes/em-dashes). Falls back to the whole file if
    fewer than two delimiter lines are found."""
    lines = template_text.splitlines()
    delimiter_indices = [i for i, line in enumerate(lines) if DELIMITER_RE.match(line)]
    if len(delimiter_indices) >= 2:
        start, end = delimiter_indices[0], delimiter_indices[1]
        return "\n".join(lines[start + 1 : end])
    return template_text


def extract_delimiter_line(template_text: str) -> str:
    """The exact delimiter line text used in the template, so the same
    style of separator can be reused between generated question blocks."""
    for line in template_text.splitlines():
        if DELIMITER_RE.match(line):
            return line
    return DEFAULT_DELIMITER_LINE


def detect_placeholder(block: str) -> str | None:
    match = PLACEHOLDER_RE.search(block)
    return match.group(1) if match else None


def _regenerate_rows(block: str, code_count: int, categories: list[dict]) -> str:
    lines = block.split("\n")
    output = []
    i = 0
    while i < len(lines):
        match = ROW_RE.match(lines[i])
        if not match:
            output.append(lines[i])
            i += 1
            continue

        indent = match.group("indent")
        group_values = []
        while i < len(lines):
            m = ROW_RE.match(lines[i])
            if not m:
                break
            group_values.append(m.group("val").strip())
            i += 1

        is_code_slots = bool(group_values) and all(CODE_SLOT_RE.match(v) for v in group_values)
        if is_code_slots:
            output.extend(f'{indent}<row label="r{n}">code{n}</row>' for n in range(1, code_count + 1))
        else:
            output.extend(
                f'{indent}<row label="r{cat["code"]}">{cat["label"]}</row>' for cat in categories
            )

    return "\n".join(output)


def build_block_xml(template_block: str, placeholder: str, new_name: str, code_count: int, categories: list[dict]) -> str:
    text = template_block.replace(placeholder, new_name)
    text = TITLE_RE.sub(f"<title>v{new_name} Coded Data</title>", text)
    text = _regenerate_rows(text, code_count, categories)
    return text.strip("\n")


def build_closed_others_block(
    template_block: str, placeholder: str, new_name: str, code_offset: int, code_count: int, categories: list[dict]
) -> str:
    """A closed+other question's .dat file holds the closed answers first,
    then the OTC-coded ("other, specify") answers -- so its XML should
    describe only the real, OTC-driven categories (matching what a regular
    question would produce), not a generic code1..codeN placeholder block
    sized to the merged file. Emits a single self-contained <checkbox> block
    (no separate loader variable) whose virtual script reads the OTC-coded
    slice of the .dat file directly, at its offset within the merged row."""
    text = template_block.replace(placeholder, new_name)
    match = CHECKBOX_BLOCK_RE.search(text)
    if not match:
        # Template doesn't have a recognizable checkbox block -- fall back to
        # the regular two-part rendering rather than producing nothing.
        return build_block_xml(template_block, placeholder, new_name, code_offset + code_count, categories)

    checkbox_block = match.group(0)
    checkbox_block = TITLE_RE.sub(f"<title>v{new_name} Coded Data</title>", checkbox_block)
    checkbox_block = _regenerate_rows(checkbox_block, code_count, categories)
    new_virtual = (
        "<virtual>\n"
        f"respData = {new_name}.get(record)\n"
        "if respData:\n"
        f"    for i in range(1, {code_count + 1}):\n"
        f'        val = respData["code" + str({code_offset} + i)].strip("\\r\\n")\n'
        "        if val:\n"
        '            data.attr("r" + val).val = 1\n'
        "  </virtual>"
    )
    checkbox_block = VIRTUAL_BLOCK_RE.sub(lambda _m: new_virtual, checkbox_block, count=1)
    exec_block = f'<exec when="virtualInit">\n{new_name} = File("{new_name}.dat","record")\n</exec>'
    return f"{exec_block}\n\n{checkbox_block}".strip("\n")


def assemble_xml(template_text: str, question_entries: list[dict]) -> tuple[bytes | None, list[str]]:
    """question_entries: [{"name": "q5a_coded", "code_count": 4, "categories": [...]}].
    An entry with "code_offset" set is a closed+other question -- rendered as
    a single real-labels-only block via build_closed_others_block instead of
    the regular two-part template."""
    warnings = []
    block = extract_template_block(template_text)
    placeholder = detect_placeholder(block)
    if not placeholder:
        warnings.append(
            "לא ניתן לזהות את שם השאלה בתבנית ה-XML (תבנית תקינה צריכה לכלול שורה כמו "
            '\'name = File("name.dat","record")\') — קובץ ה-XML לא הופק'
        )
        return None, warnings

    parts = []
    for entry in question_entries:
        if entry.get("code_offset") is not None:
            parts.append(
                build_closed_others_block(
                    block, placeholder, entry["name"], entry["code_offset"], entry["code_count"], entry["categories"]
                )
            )
        else:
            parts.append(build_block_xml(block, placeholder, entry["name"], entry["code_count"], entry["categories"]))
    delimiter = extract_delimiter_line(template_text)
    separator = f"\n\n{delimiter}\n\n"
    full_xml = (separator.join(parts) + "\n").encode("utf-8")
    return full_xml, warnings
