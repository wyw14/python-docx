"""Unit-test suite for `docx.oxml.document` module."""

from __future__ import annotations

from typing import cast

from docx.oxml.document import CT_Body
from docx.oxml.table import CT_Tbl
from docx.oxml.text.hyperlink import CT_Hyperlink
from docx.oxml.text.paragraph import CT_P

from ..unitutil.cxml import element


class DescribeCT_Body:
    """Unit-test suite for selected units of `docx.oxml.document.CT_Body`."""

    def it_counts_each_bookmark_name_by_actual_occurrence(self):
        body = cast(
            CT_Body,
            element(
                "w:body/("
                "w:p/w:bookmarkStart{w:id=1,w:name=alpha},"
                "w:p/w:bookmarkStart{w:id=2,w:name=beta},"
                "w:p/(w:bookmarkStart{w:id=3,w:name=alpha},"
                "w:bookmarkStart{w:id=4,w:name=alpha}),"
                "w:tbl/w:tr/w:tc/w:p/w:bookmarkStart{w:id=5,w:name=beta}"
                ")"
            ),
        )

        assert body.bookmark_name_counts == {"alpha": 3, "beta": 2}

    def it_iterates_hyperlinks_in_document_order_through_nested_tables(self):
        body = cast(
            CT_Body,
            element(
                "w:body/("
                "w:p/w:hyperlink{w:anchor=p1},"
                "w:p/w:hyperlink{w:anchor=p2},"
                "w:tbl/w:tr/w:tc/("
                "w:p/w:hyperlink{w:anchor=c1},"
                "w:tbl/w:tr/w:tc/w:p/w:hyperlink{w:anchor=n1},"
                "w:p/w:hyperlink{w:anchor=c2}"
                "),"
                "w:p/w:hyperlink{w:anchor=p3}"
                ")"
            ),
        )

        hyperlinks = list(body.iter_hyperlink_elements())

        assert [type(h) for h in hyperlinks] == [CT_Hyperlink] * 6
        assert [h.anchor for h in hyperlinks] == [
            "p1",
            "p2",
            "c1",
            "n1",
            "c2",
            "p3",
        ]

    def it_counts_merged_cells_by_physical_tc_occurrence(self):
        # -- one horizontally-spanning tc in the first row and a vMerge continuation
        # -- tc in the second row, each physical tc contributing its own hyperlink --
        body = cast(
            CT_Body,
            element(
                "w:body/w:tbl/("
                "w:tr/w:tc/("
                "w:tcPr/w:gridSpan{w:val=2},"
                "w:p/w:hyperlink{w:anchor=span}"
                "),"
                "w:tr/w:tc/("
                "w:tcPr/w:vMerge,"
                "w:p/w:hyperlink{w:anchor=continue}"
                ")"
                ")"
            ),
        )

        hyperlinks = list(body.iter_hyperlink_elements())

        # -- each physical w:tc is walked exactly once, unlike the layout-grid cell
        # -- proxies which would repeat the span cell and redirect the continuation --
        assert [h.anchor for h in hyperlinks] == ["span", "continue"]

    def it_knows_its_inner_content_block_item_elements(self):
        body = cast(CT_Body, element("w:body/(w:tbl, w:p,w:p)"))
        assert [type(e) for e in body.inner_content_elements] == [CT_Tbl, CT_P, CT_P]
