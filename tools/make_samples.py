"""Build the scenario documents in samples/ (developer tool, needs Word).

01  basic 12 page practice document (same generator as the in-app guide)
02  a thesis-shaped document: cover, roman numbered front matter with a table
    of contents field that spans two pages, chapters with their own headers,
    footnotes, and a landscape appendix with a long table whose header row
    repeats on every page
03  a document with Track Changes switched on and pending revisions

Run:  py -3 tools\\make_samples.py
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

import i18n  # noqa: E402
import sample_doc  # noqa: E402

OUT = ROOT / "samples"
LOREM = (
    "Penelitian ini membahas penerapan metode yang terukur pada data lapangan. Setiap tahap dicatat dengan "
    "rinci agar hasilnya dapat diulang dan diperiksa oleh pembaca. "
)


def word():
    import pythoncom
    import pywintypes
    from win32com.client import dynamic

    pythoncom.CoInitialize()
    unknown = pythoncom.CoCreateInstanceEx(pywintypes.IID("{000209FF-0000-0000-C000-000000000046}"), None,
                                           pythoncom.CLSCTX_LOCAL_SERVER, None, (pythoncom.IID_IDispatch,))[0]
    app = dynamic.Dispatch(unknown)
    app.Visible = False
    app.DisplayAlerts = 0
    return app


def append(doc, text: str, style=None, page_break_before: bool = False):
    end = doc.Content
    end.Collapse(0)
    end.InsertAfter(text)
    paragraph = doc.Paragraphs.Last.Range
    if style is not None:
        paragraph.Style = doc.Styles(style)
    paragraph.ParagraphFormat.PageBreakBefore = page_break_before
    doc.Content.InsertParagraphAfter()


def section_break(doc):
    end = doc.Content
    end.Collapse(0)
    end.InsertBreak(2)  # next page section break


def thesis(app, path: Path) -> None:
    doc = app.Documents.Add()
    doc.PageSetup.TopMargin = doc.PageSetup.BottomMargin = 72
    # Section 1: cover
    append(doc, "SKRIPSI CONTOH WORDSPLITTER", -1)
    doc.Paragraphs(1).Range.Font.Size = 20
    doc.Paragraphs(1).Alignment = 1
    for line in ("Disusun sebagai dokumen dummy", "untuk simulasi pemisahan", "Tahun 2026"):
        append(doc, line)
    section_break(doc)
    # Section 2: front matter (roman)
    append(doc, "KATA PENGANTAR", -2)
    for _ in range(6):
        append(doc, LOREM * 3)
    append(doc, "DAFTAR ISI", -2, page_break_before=True)
    append(doc, "[daftar isi]")
    section_break(doc)
    # Section 3: chapters (arabic, restart at 1)
    for number, title in enumerate(("PENDAHULUAN", "TINJAUAN PUSTAKA", "METODE PENELITIAN",
                                    "HASIL DAN PEMBAHASAN", "PENUTUP"), start=1):
        append(doc, f"BAB {number} {title}", -2, page_break_before=number > 1)
        for k in range(1, 4):
            append(doc, f"{number}.{k} Subbab {k}", -3)
            for _ in range(4):
                append(doc, LOREM * 3)
        note_range = doc.Paragraphs.Last.Range
        note_range.Collapse(0)
        note_range.Move(1, -2)
        doc.Footnotes.Add(note_range, Text=f"Catatan kaki untuk Bab {number}.")
    section_break(doc)
    # Section 4: landscape appendix with a long table
    append(doc, "LAMPIRAN TABEL DATA", -2)
    end = doc.Content
    end.Collapse(0)
    table = doc.Tables.Add(end, 90, 4)
    table.Borders.Enable = True
    for column, caption in enumerate(("No", "Responden", "Skor", "Keterangan"), start=1):
        table.Cell(1, column).Range.Text = caption
    table.Rows(1).HeadingFormat = True
    table.Rows(1).Range.Font.Bold = True
    for row in range(2, 91):
        for column, value in enumerate((str(row - 1), f"Responden {row - 1}", str(60 + (row * 7) % 40), "Valid"), start=1):
            table.Cell(row, column).Range.Text = value
    append(doc, "Akhir lampiran.")

    sections = doc.Sections
    cover, front, body, appendix = (sections(i) for i in range(1, 5))
    appendix.PageSetup.Orientation = 1
    for section, header in ((front, "Bagian Awal"), (body, "Isi Skripsi"), (appendix, "Lampiran")):
        for story in (section.Headers(1), section.Footers(1)):
            story.LinkToPrevious = False
        section.Headers(1).Range.Text = header
        section.Footers(1).PageNumbers.Add(1, True)
    cover.Footers(1).Range.Text = ""
    front.Footers(1).PageNumbers.NumberStyle = 2  # i, ii, iii
    front.Footers(1).PageNumbers.RestartNumberingAtSection = True
    front.Footers(1).PageNumbers.StartingNumber = 1
    body.Footers(1).PageNumbers.NumberStyle = 0
    body.Footers(1).PageNumbers.RestartNumberingAtSection = True
    body.Footers(1).PageNumbers.StartingNumber = 1

    placeholder = next(p.Range for p in doc.Paragraphs if p.Range.Text.strip() == "[daftar isi]")
    placeholder.MoveEnd(1, -1)  # keep the paragraph mark, replace only the text
    doc.TablesOfContents.Add(Range=placeholder, UseHeadingStyles=True, UpperHeadingLevel=1, LowerHeadingLevel=2)
    doc.TablesOfContents(1).Update()
    doc.SaveAs2(FileName=str(path), FileFormat=16)
    doc.Close(0)


def tracked(app, path: Path) -> None:
    doc = app.Documents.Add()
    for index in range(1, 9):
        append(doc, f"Bagian {index}", -2, page_break_before=index > 1)
        for _ in range(3):
            append(doc, LOREM * 4)
    doc.TrackRevisions = True
    doc.Paragraphs(2).Range.InsertBefore("[Tambahan yang dilacak] ")
    doc.Paragraphs(8).Range.Words(1).Delete()
    doc.SaveAs2(FileName=str(path), FileFormat=16)
    doc.Close(0)


def neutralise_metadata(path: Path, author: str = "WordSplitter") -> None:
    """Replace the Office user name Word stamps into authors and revisions."""
    staging = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(staging, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "docProps/core.xml":
                text = data.decode("utf-8")
                text = re.sub(r"(<dc:creator>).*?(</dc:creator>)", rf"\g<1>{author}\g<2>", text)
                text = re.sub(r"(<cp:lastModifiedBy>).*?(</cp:lastModifiedBy>)", rf"\g<1>{author}\g<2>", text)
                data = text.encode("utf-8")
            elif item.filename.startswith("word/") and item.filename.endswith(".xml"):
                name = author.encode()
                data = re.sub(rb'(w\d*:author)="[^"]*"', lambda m: m.group(1) + b'="' + name + b'"', data)
                data = re.sub(rb'(w\d*:initials)="[^"]*"', lambda m: m.group(1) + b'="WS"', data)
                data = re.sub(rb'(w15:providerId)="[^"]*"', lambda m: m.group(1) + b'="None"', data)
                data = re.sub(rb'(w15:userId)="[^"]*"', lambda m: m.group(1) + b'="' + name + b'"', data)
            target.writestr(item, data)
    staging.replace(path)


def main() -> None:
    OUT.mkdir(exist_ok=True)
    tr = i18n.Translator("id")
    sample_doc.build(OUT / "01-dasar-12-halaman.docx", tr.t, lang="id-ID")
    app = word()
    try:
        thesis(app, OUT / "02-skripsi-section-toc-tabel.docx")
        tracked(app, OUT / "03-lacak-perubahan.docx")
    finally:
        app.NormalTemplate.Saved = True
        app.Quit(0)
    for item in sorted(OUT.glob("*.docx")):
        neutralise_metadata(item)
        print(item.name, item.stat().st_size)


if __name__ == "__main__":
    main()
