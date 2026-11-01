"""Bookmark-target checking for internal hyperlinks.

An internal hyperlink (a `<w:hyperlink>` with no `r:id` and a `w:anchor`) jumps to a
bookmark located by name. The objects in this module report whether the anchor of each
such link resolves to exactly one bookmark, no bookmark, or multiple bookmarks sharing
the same name, so that broken or ambiguous links can be found before a report is
generated.
"""

from __future__ import annotations

import enum
from dataclasses import dataclass


class HyperlinkBookmarkStatus(enum.Enum):
    """Match status of an internal hyperlink anchor against document bookmarks."""

    UNIQUE = "unique"
    """Exactly one `<w:bookmarkStart>` in the body has the anchor name."""

    MISSING = "missing"
    """No `<w:bookmarkStart>` in the body has the anchor name; the link has no target."""

    DUPLICATE = "duplicate"
    """More than one `<w:bookmarkStart>` has the anchor name; the jump target is ambiguous."""


@dataclass(frozen=True)
class HyperlinkBookmarkCheck:
    """Result of checking one internal hyperlink against bookmarks in the document body.

    `text` is the visible text of the hyperlink. `anchor` is the bookmark name from
    its `w:anchor` attribute. `status` reports how many body bookmarks matched:
    `~HyperlinkBookmarkStatus.UNIQUE` for exactly one, `~HyperlinkBookmarkStatus.MISSING`
    for none, and `~HyperlinkBookmarkStatus.DUPLICATE` for two or more. `bookmark_count`
    is the number of matching `<w:bookmarkStart>` elements (1, 0, or 2+ respectively).
    """

    text: str
    anchor: str
    status: HyperlinkBookmarkStatus
    bookmark_count: int
