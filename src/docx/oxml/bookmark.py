"""Custom element classes related to bookmarks."""

from __future__ import annotations

from docx.oxml.simpletypes import ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, RequiredAttribute


class CT_BookmarkStart(BaseOxmlElement):
    """`<w:bookmarkStart>` element, marking the start of a bookmarked range.

    The bookmark `name` is the value referenced by the `w:anchor` attribute of an
    internal `<w:hyperlink>`. Bookmark names are not guaranteed to be unique within a
    document, so more than one bookmark start may carry the same name.
    """

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    name: str = RequiredAttribute("w:name", ST_String)  # pyright: ignore[reportAssignmentType]
