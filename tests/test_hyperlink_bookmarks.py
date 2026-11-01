# pyright: reportPrivateUsage=false

"""Integration-test suite for `Document.check_internal_hyperlinks()`.

These tests build real documents using the python-docx object graph (with hyperlink
and bookmark elements inserted as WordprocessingML), exercise body paragraphs,
nested tables, merged cells, duplicate bookmark names, missing targets, external
links and field-code links, and verify the results survive a save/reload cycle.
"""

from __future__ import annotations

from dataclasses import astuple
from io import BytesIO
from typing import cast

import pytest

from docx import Document
from docx.bookmark import HyperlinkBookmarkCheck, HyperlinkBookmarkStatus
from docx.document import Document as DocumentObject
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.oxml.text.hyperlink import CT_Hyperlink
from docx.table import _Cell

# -- shorthands ---------------------------------------------------------------------------

UNIQUE = HyperlinkBookmarkStatus.UNIQUE
MISSING = HyperlinkBookmarkStatus.MISSING
DUPLICATE = HyperlinkBookmarkStatus.DUPLICATE


def _hyperlink(anchor: str | None, text: str, rId: str | None = None) -> CT_Hyperlink:
    """A `w:hyperlink` containing a single run with `text`."""
    rId_attr = f' r:id="{rId}"' if rId else ""
    anchor_attr = f' w:anchor="{anchor}"' if anchor is not None else ""
    return cast(
        CT_Hyperlink,
        parse_xml(
            f"<w:hyperlink {nsdecls('w', 'r')}{rId_attr}{anchor_attr}>"
            f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
            f"</w:hyperlink>"
        ),
    )


def _two_run_hyperlink(anchor: str, first: str, second: str) -> CT_Hyperlink:
    """A `w:hyperlink` whose visible text spans two runs."""
    return cast(
        CT_Hyperlink,
        parse_xml(
            f'<w:hyperlink {nsdecls("w")} w:anchor="{anchor}">'
            f'<w:r><w:t xml:space="preserve">{first}</w:t></w:r>'
            f'<w:r><w:t xml:space="preserve">{second}</w:t></w:r>'
            f"</w:hyperlink>"
        ),
    )


def _bookmark_start(bookmark_id: int, name: str):
    """A `w:bookmarkStart` element."""
    return parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="{bookmark_id}" w:name="{name}"/>')


@pytest.fixture
def document() -> DocumentObject:
    """A document exercising each internal-hyperlink bookmark-target case."""
    document = Document()

    # -- body paragraphs with internal links in document order --
    document.add_paragraph()._p.append(_hyperlink("bm_unique", "unique link"))
    document.add_paragraph()._p.append(_two_run_hyperlink("bm_unique", "multi ", "run link"))
    document.add_paragraph()._p.append(_hyperlink("bm_missing", "missing link"))
    document.add_paragraph()._p.append(_hyperlink("bm_dup", "duplicate link"))

    # -- external links are not internal links, even when an anchor or a URL fragment
    # -- is present --
    rId = document.part.rels.get_or_add_ext_rel(RT.HYPERLINK, "https://example.com/page#bm_unique")
    document.add_paragraph()._p.append(
        _hyperlink("bm_unique", "external link with anchor", rId=rId)
    )
    plain_rId = document.part.rels.get_or_add_ext_rel(RT.HYPERLINK, "https://example.com/")
    document.add_paragraph()._p.append(_hyperlink(None, "plain external link", rId=plain_rId))

    # -- a hyperlink with neither an address nor an anchor is not checked --
    document.add_paragraph()._p.append(_hyperlink(None, "anchorless link"))

    # -- bookmarks in body paragraphs, including one duplicated name --
    document.add_paragraph()._p.append(_bookmark_start(0, "bm_unique"))
    document.add_paragraph()._p.append(_bookmark_start(1, "bm_dup"))
    document.add_paragraph()._p.append(_bookmark_start(2, "bm_dup"))

    # -- a table with a nested table, plus horizontal and vertical merges --
    table = document.add_table(rows=4, cols=2)

    cell_0_0 = table.cell(0, 0)
    cell_0_0.paragraphs[0]._p.append(_bookmark_start(3, "bm_cell"))
    cell_0_0.paragraphs[0]._p.append(_hyperlink("bm_cell", "table link"))

    nested_table = table.cell(0, 1).add_table(1, 2)
    nested_cell = nested_table.cell(0, 0)
    nested_cell.paragraphs[0]._p.append(_bookmark_start(4, "bm_nested"))
    nested_cell.paragraphs[0]._p.append(_hyperlink("bm_nested", "nested link"))
    nested_table.cell(0, 1).paragraphs[0]._p.append(_hyperlink("bm_missing", "nested missing"))

    # -- a horizontally merged cell (one physical w:tc spanning two grid columns) --
    merged_h = table.cell(1, 0).merge(table.cell(1, 1))
    merged_h.paragraphs[0]._p.append(_hyperlink("bm_cell", "horizontal merge link"))

    # -- a vertically merged cell: restart tc on row 2, continuation tc on row 3 --
    table.cell(2, 0).merge(table.cell(3, 0))
    table.cell(2, 0).paragraphs[0]._p.append(_hyperlink("bm_cell", "vertical root link"))
    # -- the continuation cell is itself a physical w:tc element; content it carries
    # -- must be counted once rather than redirected to the root cell by a grid proxy --
    continue_tc = table._tbl.tr_lst[3].tc_lst[0]
    _Cell(continue_tc, table).add_paragraph()._p.append(
        _hyperlink("bm_missing", "vertical continue link")
    )

    # -- a hyperlink produced by a Word field code is not a w:hyperlink element --
    document.add_paragraph()._p.append(
        parse_xml(
            f'<w:fldSimple {nsdecls("w")} w:instr=" HYPERLINK \\l &quot;bm_missing&quot; ">'
            "<w:r><w:t>field-code link</w:t></w:r>"
            "</w:fldSimple>"
        )
    )

    return document


class DescribeInternalHyperlinkBookmarkChecking:
    """Integration tests for bookmark-target checking of internal hyperlinks."""

    def it_reports_every_internal_link_status_in_document_order(self, document: DocumentObject):
        checks = document.check_internal_hyperlinks()

        assert checks == (
            HyperlinkBookmarkCheck("unique link", "bm_unique", UNIQUE, 1),
            HyperlinkBookmarkCheck("multi run link", "bm_unique", UNIQUE, 1),
            HyperlinkBookmarkCheck("missing link", "bm_missing", MISSING, 0),
            HyperlinkBookmarkCheck("duplicate link", "bm_dup", DUPLICATE, 2),
            HyperlinkBookmarkCheck("table link", "bm_cell", UNIQUE, 1),
            HyperlinkBookmarkCheck("nested link", "bm_nested", UNIQUE, 1),
            HyperlinkBookmarkCheck("nested missing", "bm_missing", MISSING, 0),
            HyperlinkBookmarkCheck("horizontal merge link", "bm_cell", UNIQUE, 1),
            HyperlinkBookmarkCheck("vertical root link", "bm_cell", UNIQUE, 1),
            HyperlinkBookmarkCheck("vertical continue link", "bm_missing", MISSING, 0),
        )

    def it_omits_external_anchorless_and_field_code_links(self, document: DocumentObject):
        checked_texts = {check.text for check in document.check_internal_hyperlinks()}

        assert "external link with anchor" not in checked_texts
        assert "plain external link" not in checked_texts
        assert "anchorless link" not in checked_texts
        assert "field-code link" not in checked_texts

    def it_does_not_modify_hyperlinks_bookmarks_or_relationships(self, document: DocumentObject):
        part_xml_before = document.part.rels.xml
        document_xml_before = document.element.xml

        document.check_internal_hyperlinks()

        assert document.element.xml == document_xml_before
        assert document.part.rels.xml == part_xml_before

    def it_gives_the_same_results_after_save_and_reload(self, document: DocumentObject):
        before = [astuple(check) for check in document.check_internal_hyperlinks()]

        stream = BytesIO()
        document.save(stream)
        stream.seek(0)
        reloaded = Document(stream)
        after = [astuple(check) for check in reloaded.check_internal_hyperlinks()]

        assert after == before
