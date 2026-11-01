"""Step implementations for internal-hyperlink bookmark-target checking features."""

from __future__ import annotations

from dataclasses import astuple
from io import BytesIO
from typing import cast

from behave import given, then, when
from behave.runner import Context

from docx import Document
from docx.bookmark import HyperlinkBookmarkCheck, HyperlinkBookmarkStatus
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.oxml.text.hyperlink import CT_Hyperlink
from docx.table import _Cell

# -- builders -----------------------------------------------------------------------------


def _hyperlink(anchor: str | None, text: str, rId: str | None = None) -> CT_Hyperlink:
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


def _bookmark_start(bookmark_id: int, name: str):
    return parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="{bookmark_id}" w:name="{name}"/>')


def _body_document():
    """A document of body paragraphs exercising unique/missing/duplicate targets."""
    document = Document()

    document.add_paragraph()._p.append(_hyperlink("bm_unique", "unique link"))
    document.add_paragraph()._p.append(_hyperlink("bm_missing", "missing link"))
    document.add_paragraph()._p.append(_hyperlink("bm_dup", "duplicate link"))

    # -- external links must never be treated as internal bookmark jumps, whether or
    # -- not the address contains a "#" fragment or the anchor attribute is present --
    rId = document.part.rels.get_or_add_ext_rel(RT.HYPERLINK, "https://example.com/page#bm_unique")
    document.add_paragraph()._p.append(
        _hyperlink("bm_unique", "external link with anchor", rId=rId)
    )
    plain_rId = document.part.rels.get_or_add_ext_rel(RT.HYPERLINK, "https://example.com/")
    document.add_paragraph()._p.append(_hyperlink(None, "plain external link", rId=plain_rId))

    # -- hyperlinks with no anchor are out of scope --
    document.add_paragraph()._p.append(_hyperlink(None, "anchorless link"))

    document.add_paragraph()._p.append(_bookmark_start(0, "bm_unique"))
    document.add_paragraph()._p.append(_bookmark_start(1, "bm_dup"))
    document.add_paragraph()._p.append(_bookmark_start(2, "bm_dup"))

    # -- a field-code hyperlink is not a w:hyperlink element and is not detected --
    document.add_paragraph()._p.append(
        parse_xml(
            f'<w:fldSimple {nsdecls("w")} w:instr=" HYPERLINK \\l &quot;bm_missing&quot; ">'
            "<w:r><w:t>field-code link</w:t></w:r>"
            "</w:fldSimple>"
        )
    )

    return document


def _add_table_with_links(document):
    """Add a table containing nested-table and merged-cell internal hyperlinks."""
    table = document.add_table(rows=4, cols=2)

    table.cell(0, 0).paragraphs[0]._p.append(_bookmark_start(3, "bm_cell"))
    table.cell(0, 0).paragraphs[0]._p.append(_hyperlink("bm_cell", "table link"))

    nested_table = table.cell(0, 1).add_table(1, 2)
    nested_table.cell(0, 0).paragraphs[0]._p.append(_bookmark_start(4, "bm_nested"))
    nested_table.cell(0, 0).paragraphs[0]._p.append(_hyperlink("bm_nested", "nested link"))
    nested_table.cell(0, 1).paragraphs[0]._p.append(_hyperlink("bm_missing", "nested missing"))

    merged_h = table.cell(1, 0).merge(table.cell(1, 1))
    merged_h.paragraphs[0]._p.append(_hyperlink("bm_cell", "horizontal merge link"))

    table.cell(2, 0).merge(table.cell(3, 0))
    table.cell(2, 0).paragraphs[0]._p.append(_hyperlink("bm_cell", "vertical root link"))
    continue_tc = table._tbl.tr_lst[3].tc_lst[0]
    _Cell(continue_tc, table).add_paragraph()._p.append(
        _hyperlink("bm_missing", "vertical continue link")
    )


# -- expected values ----------------------------------------------------------------------

EXPECTED_BODY_CHECKS = (
    HyperlinkBookmarkCheck("unique link", "bm_unique", HyperlinkBookmarkStatus.UNIQUE, 1),
    HyperlinkBookmarkCheck("missing link", "bm_missing", HyperlinkBookmarkStatus.MISSING, 0),
    HyperlinkBookmarkCheck("duplicate link", "bm_dup", HyperlinkBookmarkStatus.DUPLICATE, 2),
)

EXPECTED_TABLE_CHECK_TEXTS = (
    "table link",
    "nested link",
    "nested missing",
    "horizontal merge link",
    "vertical root link",
    "vertical continue link",
)

EXCLUDED_TEXTS = (
    "external link with anchor",
    "plain external link",
    "anchorless link",
    "field-code link",
)


# given ===================================================================================


@given("a document containing bookmarks and internal hyperlinks")
def given_document_containing_bookmarks_and_internal_hyperlinks(context: Context):
    context.document = _body_document()


@given("a document containing internal hyperlinks in nested and merged table cells")
def given_document_with_nested_and_merged_table_links(context: Context):
    context.document = Document()
    _add_table_with_links(context.document)


@given("a checked document containing bookmarks and internal hyperlinks")
def given_a_checked_document(context: Context):
    document = _body_document()
    _add_table_with_links(document)
    context.document = document
    context.checks_before = document.check_internal_hyperlinks()
    context.document_xml_before = document.element.xml
    context.rels_xml_before = document.part.rels.xml


# when ====================================================================================


@when("I check the internal hyperlink bookmark targets")
def when_check_internal_hyperlink_bookmark_targets(context: Context):
    context.checks = context.document.check_internal_hyperlinks()


@when("I save and reload the document and check it again")
def when_save_reload_and_check_again(context: Context):
    stream = BytesIO()
    context.document.save(stream)
    stream.seek(0)
    context.document = Document(stream)
    context.checks_after = context.document.check_internal_hyperlinks()


# then ====================================================================================


@then("the report lists each internal link in document order with its text, anchor, and status")
def then_report_lists_each_internal_link(context: Context):
    assert tuple(context.checks) == EXPECTED_BODY_CHECKS


@then("the duplicate-target result gives the number of matching bookmarks")
def then_duplicate_result_gives_match_count(context: Context):
    duplicate_checks = [c for c in context.checks if c.status is HyperlinkBookmarkStatus.DUPLICATE]
    assert len(duplicate_checks) == 1
    assert duplicate_checks[0].anchor == "bm_dup"
    assert duplicate_checks[0].bookmark_count == 2


@then("external links, anchorless links, and field-code links are not reported")
def then_excluded_links_are_not_reported(context: Context):
    reported_texts = [c.text for c in context.checks]
    for text in EXCLUDED_TEXTS:
        assert text not in reported_texts, f"{text!r} was unexpectedly reported"


@then("each physical table cell contributes its links exactly once in document order")
def then_each_physical_cell_contributes_links_once(context: Context):
    texts = [c.text for c in context.checks]
    assert texts == list(EXPECTED_TABLE_CHECK_TEXTS)
    for text in EXPECTED_TABLE_CHECK_TEXTS:
        assert texts.count(text) == 1, f"{text!r} was counted more than once"


@then("the reloaded document reports the same bookmark-target results")
def then_reloaded_document_reports_same_results(context: Context):
    before = [astuple(c) for c in context.checks_before]
    after = [astuple(c) for c in context.checks_after]
    assert after == before


@then("the hyperlinks, bookmarks, and relationships are unchanged by the checks")
def then_hyperlinks_bookmarks_and_relationships_unchanged(context: Context):
    assert context.document.element.xml == context.document_xml_before
    assert context.document.part.rels.xml == context.rels_xml_before
