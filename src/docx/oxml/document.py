"""Custom element classes that correspond to the document part, e.g. <w:document>."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Dict, Iterable, Iterator, List, cast

from docx.oxml.section import CT_SectPr
from docx.oxml.table import CT_Tbl
from docx.oxml.text.hyperlink import CT_Hyperlink
from docx.oxml.text.paragraph import CT_P
from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrMore, ZeroOrOne

if TYPE_CHECKING:
    from docx.oxml.bookmark import CT_BookmarkStart


class CT_Document(BaseOxmlElement):
    """``<w:document>`` element, the root element of a document.xml file."""

    body: CT_Body = ZeroOrOne("w:body")  # pyright: ignore[reportAssignmentType]

    @property
    def sectPr_lst(self) -> List[CT_SectPr]:
        """All `w:sectPr` elements directly accessible from document element.

        Note this does not include a `sectPr` child in a paragraphs wrapped in
        revision marks or other intervening layer, perhaps `w:sdt` or customXml
        elements.

        `w:sectPr` elements appear in document order. The last one is always
        `w:body/w:sectPr`, all preceding are `w:p/w:pPr/w:sectPr`.
        """
        xpath = "./w:body/w:p/w:pPr/w:sectPr | ./w:body/w:sectPr"
        return self.xpath(xpath)


class CT_Body(BaseOxmlElement):
    """`w:body`, the container element for the main document story in `document.xml`."""

    add_p: Callable[[], CT_P]
    get_or_add_sectPr: Callable[[], CT_SectPr]
    p_lst: List[CT_P]
    tbl_lst: List[CT_Tbl]

    _insert_tbl: Callable[[CT_Tbl], CT_Tbl]

    p = ZeroOrMore("w:p", successors=("w:sectPr",))
    tbl = ZeroOrMore("w:tbl", successors=("w:sectPr",))
    sectPr: CT_SectPr | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:sectPr", successors=()
    )

    def add_section_break(self) -> CT_SectPr:
        """Return `w:sectPr` element for new section added at end of document.

        The last `w:sectPr` becomes the second-to-last, with the new `w:sectPr` being an
        exact clone of the previous one, except that all header and footer references
        are removed (and are therefore now "inherited" from the prior section).

        A copy of the previously-last `w:sectPr` will now appear in a new `w:p` at the
        end of the document. The returned `w:sectPr` is the sentinel `w:sectPr` for the
        document (and as implemented, `is` the prior sentinel `w:sectPr` with headers
        and footers removed).
        """
        # ---get the sectPr at file-end, which controls last section (sections[-1])---
        sentinel_sectPr = self.get_or_add_sectPr()
        # ---add exact copy to new `w:p` element; that is now second-to last section---
        self.add_p().set_sectPr(sentinel_sectPr.clone())
        # ---remove any header or footer references from "new" last section---
        for hdrftr_ref in sentinel_sectPr.xpath("w:headerReference|w:footerReference"):
            sentinel_sectPr.remove(hdrftr_ref)
        # ---the sentinel `w:sectPr` now controls the new last section---
        return sentinel_sectPr

    def clear_content(self):
        """Remove all content child elements from this <w:body> element.

        Leave the <w:sectPr> element if it is present.
        """
        for content_elm in self.xpath("./*[not(self::w:sectPr)]"):
            self.remove(content_elm)

    @property
    def bookmark_name_counts(self) -> Dict[str, int]:
        """Mapping of each bookmark name to its number of occurrences in this body.

        Counts every `<w:bookmarkStart>` descendant of the body, including those
        nested in table cells, in document order. A name maps to 1 for an ordinary
        bookmark and to a value greater than 1 when bookmark names are duplicated.
        Names without a bookmark start simply do not appear in the mapping.
        """
        bookmark_starts = cast("list[CT_BookmarkStart]", self.xpath(".//w:bookmarkStart"))
        counts: Dict[str, int] = {}
        for bookmarkStart in bookmark_starts:
            counts[bookmarkStart.name] = counts.get(bookmarkStart.name, 0) + 1
        return counts

    @property
    def inner_content_elements(self) -> List[CT_P | CT_Tbl]:
        """Generate all `w:p` and `w:tbl` elements in this document-body.

        Elements appear in document order. Elements shaded by nesting in a `w:ins` or
        other "wrapper" element will not be included.
        """
        return self.xpath("./w:p | ./w:tbl")

    def iter_hyperlink_elements(self) -> Iterator[CT_Hyperlink]:
        """Generate each `<w:hyperlink>` in this body in document order.

        Hyperlinks in body paragraphs are generated first, along with those inside
        tables, including tables nested in cells to any depth. The physical XML
        structure is walked, so each `<w:tc>` element is visited exactly once,
        including horizontally-merged (`w:gridSpan`) and vertically-merged
        continuation (`w:vMerge="continue"`) cells; layout-grid cell proxies that
        repeat the same cell reference are deliberately not used.

        Only hyperlinks that are direct children of a `<w:p>` are generated. Links
        produced by Word field codes (`w:fldSimple` or `HYPERLINK` instrText) are not
        `<w:hyperlink>` elements and so are not generated.
        """

        def iter_in_blocks(blocks: Iterable[CT_P | CT_Tbl]) -> Iterator[CT_Hyperlink]:
            for block in blocks:
                if isinstance(block, CT_P):
                    yield from block.hyperlink_lst
                else:
                    # -- a CT_Tbl: walk each physical row/cell once, recursing into
                    # -- paragraphs and nested tables within the cell in document order --
                    for tr in block.tr_lst:
                        for tc in tr.tc_lst:
                            yield from iter_in_blocks(tc.inner_content_elements)

        yield from iter_in_blocks(self.inner_content_elements)
