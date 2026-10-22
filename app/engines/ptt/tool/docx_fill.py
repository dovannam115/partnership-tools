"""Helpers to edit .docx text without losing run formatting."""
import re


def _dedup_cells(row):
    seen, out = set(), []
    for c in row.cells:
        if id(c._tc) not in seen:
            seen.add(id(c._tc))
            out.append(c)
    return out


def iter_paragraphs(container):
    for p in container.paragraphs:
        yield p
    for t in getattr(container, "tables", []):
        for r in t.rows:
            for c in _dedup_cells(r):
                yield from iter_paragraphs(c)


def replace_in_paragraph(p, pattern, repl, count=1):
    """Replace regex `pattern` inside a paragraph even when the match is split
    across several runs. Returns the list of runs holding the new text."""
    hits = []
    for _ in range(count):
        runs = p.runs
        if not runs:
            break
        text = "".join(r.text for r in runs)
        m = re.search(pattern, text)
        if not m:
            break
        start, end = m.span()
        new = m.expand(repl) if isinstance(repl, str) else repl(m)
        pos, touched = 0, []
        for r in runs:
            rs, re_ = pos, pos + len(r.text)
            if re_ > start and rs < end:
                touched.append((r, rs, re_))
            pos = re_
        first, last = touched[0], touched[-1]
        head = first[0].text[: start - first[1]]
        tail = last[0].text[end - last[1] :]
        if first is last:
            first[0].text = head + new + tail
        else:
            first[0].text = head + new
            for r, _s, _e in touched[1:-1]:
                r.text = ""
            last[0].text = tail
        hits.append(first[0])
    return hits


def replace_everywhere(doc, pattern, repl, count=99):
    n = 0
    for p in iter_paragraphs(doc):
        n += len(replace_in_paragraph(p, pattern, repl, count))
    return n


def cell_paragraph(cell):
    """First paragraph of a cell that has any run, else the first one."""
    for p in cell.paragraphs:
        if p.runs:
            return p
    return cell.paragraphs[0]
