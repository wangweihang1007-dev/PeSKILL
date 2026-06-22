from __future__ import annotations

import re
from pathlib import Path


OUTPUT_DIR_NAME = "AI会议纪要输出"
IGNORE_TERMS = ("会议笔记", "手写", "批注")
TRANSCRIPT_EXCLUDES = ("会议纪要", "qa", "q&a", "总结", "投资可行性", "项目表")
SPEAKER_RE = re.compile(
    r"^(?:\[?\d{1,2}:\d{2}(?::\d{2})?\]?\s*)?(?:发言人\s*\d+|说话人\s*\d+|speaker\s*\d+|主持人|投资方|公司方|问|答|q|a)\s*[：:]",
    re.IGNORECASE,
)


def hidden_or_ignored(path: Path, root: Path) -> bool:
    try:
        parts = path.relative_to(root).parts
    except ValueError:
        parts = path.parts
    if any(part.startswith(".") or part.startswith("~$") for part in parts):
        return True
    if OUTPUT_DIR_NAME in parts:
        return True
    return any(term in path.name for term in IGNORE_TERMS)


def normalize_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = "\n".join(line.rstrip() for line in text.splitlines())
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def split_complete_turns(text: str, limit: int = 4500) -> list[dict]:
    lines = text.splitlines()
    turns: list[tuple[int, int, str]] = []
    start = 1
    current: list[str] = []
    for line_no, line in enumerate(lines, 1):
        if current and SPEAKER_RE.match(line.strip()):
            turns.append((start, line_no - 1, "\n".join(current).strip()))
            start = line_no
            current = [line]
        else:
            if not current:
                start = line_no
            current.append(line)
    if current:
        turns.append((start, len(lines), "\n".join(current).strip()))

    chunks: list[dict] = []
    buffer: list[str] = []
    chunk_start = 1
    chunk_end = 1
    for turn_start, turn_end, turn in turns:
        pieces = _split_long_turn(turn, limit)
        for piece in pieces:
            proposed = "\n".join(buffer + [piece]).strip()
            if buffer and len(proposed) > limit:
                chunks.append(_chunk(len(chunks) + 1, buffer, chunk_start, chunk_end))
                buffer = []
                chunk_start = turn_start
            if not buffer:
                chunk_start = turn_start
            buffer.append(piece)
            chunk_end = turn_end
    if buffer:
        chunks.append(_chunk(len(chunks) + 1, buffer, chunk_start, chunk_end))
    return chunks


def _split_long_turn(turn: str, limit: int) -> list[str]:
    if len(turn) <= limit:
        return [turn]
    sentences = re.split(r"(?<=[。！？!?；;])", turn)
    pieces: list[str] = []
    buf = ""
    for sentence in sentences:
        if buf and len(buf) + len(sentence) > limit:
            pieces.append(buf.strip())
            buf = ""
        while len(sentence) > limit:
            pieces.append(sentence[:limit].strip())
            sentence = sentence[limit:]
        buf += sentence
    if buf.strip():
        pieces.append(buf.strip())
    return pieces


def _chunk(index: int, pieces: list[str], start: int, end: int) -> dict:
    text = "\n".join(pieces).strip()
    return {
        "index": index,
        "text": text,
        "characters": len(text),
        "source_start_line": start,
        "source_end_line": end,
    }
