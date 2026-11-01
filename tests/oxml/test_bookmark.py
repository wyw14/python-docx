"""Unit-test suite for `docx.oxml.bookmark` module."""

from __future__ import annotations

from typing import cast

from docx.oxml.bookmark import CT_BookmarkStart

from ..unitutil.cxml import element


class DescribeCT_BookmarkStart:
    """Unit-test suite for `docx.oxml.bookmark.CT_BookmarkStart`."""

    def it_knows_its_bookmark_id_and_name(self):
        bookmarkStart = cast(CT_BookmarkStart, element("w:bookmarkStart{w:id=7,w:name=_Toc12345}"))

        assert bookmarkStart.id == 7
        assert bookmarkStart.name == "_Toc12345"
