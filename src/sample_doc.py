"""Generator for the practice ("dummy") document used by the in-app guide.

The file is written as raw Office Open XML with the standard library only, so
it costs nothing at start up and needs no Word. Every page after the first
starts with a heading that has "page break before" set, which makes the page
count deterministic (12 pages on any machine) without manual break characters.
The text is produced in the interface language, right-to-left where needed.
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Callable
from xml.sax.saxutils import escape

PAGES = 12
SPLIT_A = 6
SPLIT_B = 7

_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/docProps/core.xml" '
    'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    "</Types>"
)
_ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="word/document.xml"/>'
    '<Relationship Id="rId2" '
    'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
    'Target="docProps/core.xml"/>'
    "</Relationships>"
)
_DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
    'Target="styles.xml"/>'
    "</Relationships>"
)
W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _styles(font: str, rtl: bool, lang: str) -> str:
    bidi = "<w:bidi/>" if rtl else ""
    rtl_run = "<w:rtl/>" if rtl else ""
    fonts = f'<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="{escape(font)}" w:cs="{escape(font)}"/>'
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:styles {W_NS}>"
        "<w:docDefaults>"
        f'<w:rPrDefault><w:rPr>{fonts}<w:sz w:val="24"/><w:szCs w:val="24"/>'
        f'<w:lang w:val="{lang}" w:eastAsia="{lang}" w:bidi="{lang}"/>{rtl_run}</w:rPr></w:rPrDefault>'
        f'<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="300" w:lineRule="auto"/>{bidi}</w:pPr></w:pPrDefault>'
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/>'
        '<w:pPr><w:spacing w:after="120"/></w:pPr>'
        '<w:rPr><w:color w:val="595959"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>'
        '<w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="240"/>'
        '<w:outlineLvl w:val="0"/></w:pPr>'
        '<w:rPr><w:b/><w:bCs/><w:color w:val="185ABD"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>'
        "</w:styles>"
    )


def _paragraph(text: str, style: str = "", color: str = "", bold: bool = False, page_break: bool = False) -> str:
    ppr = ""
    if style or page_break:
        ppr = "<w:pPr>" + (f'<w:pStyle w:val="{style}"/>' if style else "")
        ppr += "<w:pageBreakBefore/>" if page_break else ""
        ppr += "</w:pPr>"
    rpr = ""
    if color or bold:
        rpr = "<w:rPr>" + ("<w:b/><w:bCs/>" if bold else "") + (f'<w:color w:val="{color}"/>' if color else "") + "</w:rPr>"
    return f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def build(path: Path, t: Callable[..., str], font: str = "Calibri", rtl: bool = False, lang: str = "en-US") -> Path:
    """Write the 12 page practice document to `path` and return it."""
    body = []
    for page in range(1, PAGES + 1):
        if page == 1:
            body.append(_paragraph(t("sample.title"), style="Title"))
        body.append(_paragraph(t("sample.page_heading", n=page, total=PAGES), style="Heading1", page_break=page > 1))
        body.append(_paragraph(t("sample.page_body", n=page)))
        first = page <= SPLIT_A
        body.append(_paragraph(
            t("sample.goes_file1" if first else "sample.goes_file2"),
            color="185ABD" if first else "0E7C66",
            bold=True,
        ))
    section = (
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>'
        + ("<w:bidi/>" if rtl else "")
        + "</w:sectPr>"
    )
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document {W_NS}><w:body>{"".join(body)}{section}</w:body></w:document>'
    )
    core = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/">'
        f"<dc:title>{escape(t('sample.title'))}</dc:title><dc:creator>WordSplitter</dc:creator>"
        "</cp:coreProperties>"
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    staging = path.with_name(path.name + ".tmp")
    with zipfile.ZipFile(staging, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _CONTENT_TYPES)
        archive.writestr("_rels/.rels", _ROOT_RELS)
        archive.writestr("word/_rels/document.xml.rels", _DOC_RELS)
        archive.writestr("word/document.xml", document)
        archive.writestr("word/styles.xml", _styles(font, rtl, lang))
        archive.writestr("docProps/core.xml", core)
    staging.replace(path)
    return path
