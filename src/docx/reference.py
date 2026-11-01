# pyright: reportPrivateUsage=false

"""Inventory of references made to embedded image parts.

The objects in this module distinguish the number of *image parts* in a package from
the number of times images from those parts actually *appear* in the document. A
single image part can be referenced by many inline pictures, possibly from more than
one story part (for example the document body and a header part).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, List, Sequence, Tuple, cast, overload

from docx.exceptions import ImageReferenceError
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.parts.image import ImagePart

if TYPE_CHECKING:
    from docx.oxml.shape import CT_Blip
    from docx.parts.document import DocumentPart
    from docx.parts.story import StoryPart


# -- XPath selecting the blip of every inline picture in a story, in document order.
# -- Restricting the path to `wp:inline` excludes floating (`wp:anchor`) pictures and
# -- requiring `pic:pic` within `a:graphicData` excludes charts and SmartArt, which
# -- use other graphicData content. Leading `.//` matches pictures at any depth,
# -- including those in table cells and nested tables.
_INLINE_PICTURE_BLIP_XPATH = (
    ".//wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip"
)


class ImageReference:
    """A single occurrence of an embedded inline picture within a story part.

    An occurrence corresponds to one ``a:blip`` element having an ``r:embed``
    relationship to an image part. The same image part referenced from multiple
    locations produces multiple |ImageReference| objects, one for each occurrence.
    """

    def __init__(self, partname: str, index: int):
        self._partname = partname
        self._index = index

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, ImageReference):
            return NotImplemented
        return self._partname == other._partname and self._index == other._index

    def __hash__(self) -> int:
        return hash((self._partname, self._index))

    def __repr__(self) -> str:
        return "ImageReference(partname='%s', index=%d)" % (self._partname, self._index)

    @property
    def index(self) -> int:
        """1-based sequence number of this occurrence within its containing story part.

        Numbering restarts at 1 for each story part and counts only embedded inline
        pictures, in document order.
        """
        return self._index

    @property
    def partname(self) -> str:
        """Partname of the story part containing this occurrence, e.g.
        ``'/word/document.xml'`` or ``'/word/header1.xml'``."""
        return self._partname


class ImageReferenceGroup:
    """All occurrences that reference a single image part.

    A group is produced for each distinct referenced |ImagePart|. Two image parts
    having identical bytes are reported as separate groups; grouping is by part, not
    by image content.
    """

    def __init__(self, image_part: ImagePart, references: Tuple[ImageReference, ...]):
        self._image_part = image_part
        self._references = references

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, ImageReferenceGroup):
            return NotImplemented
        return (
            self._image_part.partname == other._image_part.partname
            and self.content_type == other.content_type
            and self._references == other._references
        )

    def __hash__(self) -> int:
        return hash((self.partname, self.content_type, self._references))

    def __iter__(self) -> Iterator[ImageReference]:
        return iter(self._references)

    def __len__(self) -> int:
        return len(self._references)

    def __repr__(self) -> str:
        return (
            "ImageReferenceGroup(partname='%s', content_type='%s', count=%d)"
            % (self.partname, self.content_type, len(self._references))
        )

    @property
    def content_type(self) -> str:
        """Content type of the image part, e.g. ``'image/png'``."""
        return self._image_part.content_type

    @property
    def count(self) -> int:
        """Total number of inline-picture occurrences referencing this image part."""
        return len(self._references)

    @property
    def partname(self) -> str:
        """Partname of the image part, e.g. ``'/word/media/image1.png'``."""
        return str(self._image_part.partname)

    @property
    def reference_count(self) -> int:
        """Alias of :attr:`count`."""
        return len(self._references)

    @property
    def references(self) -> Tuple[ImageReference, ...]:
        """The occurrences of this image part, in traversal (document body then
        header/footer parts) and within-part document order."""
        return self._references


class ImageReferences(Sequence[ImageReferenceGroup]):
    """Sequence of |ImageReferenceGroup| objects, one per referenced image part.

    Scans the main document story and the existing header and footer parts for
    embedded inline pictures, including those in tables and nested tables. The scan
    is strictly read-only; it does not alter document XML, relationships, parts, or
    image bytes. Externally-linked pictures (those having only ``r:link`` and no
    ``r:embed``) are ignored and never fetched. Charts and SmartArt are not pictures
    and are ignored.

    An |ImageReferenceError| is raised when a picture's ``r:embed`` relationship is
    missing or does not target an image part, rather than silently producing a
    fabricated result.
    """

    def __init__(self, document_part: DocumentPart):
        super(ImageReferences, self).__init__()
        self._document_part = document_part

    @overload
    def __getitem__(self, index: int) -> ImageReferenceGroup: ...

    @overload
    def __getitem__(self, index: slice) -> Tuple[ImageReferenceGroup, ...]: ...

    def __getitem__(
        self, index: int | slice
    ) -> ImageReferenceGroup | Tuple[ImageReferenceGroup, ...]:
        return self._groups[index]

    def __iter__(self) -> Iterator[ImageReferenceGroup]:
        return iter(self._groups)

    def __len__(self) -> int:
        return len(self._groups)

    @property
    def _groups(self) -> Tuple[ImageReferenceGroup, ...]:
        """The |ImageReferenceGroup| objects for this document, in first-reference
        order.

        Computed on each access so the collection reflects the current state of the
        document, like |InlineShapes| does.
        """
        # -- map each referenced image part to its occurrences, recording the order
        # -- in which image parts are first encountered --
        references_by_image_part: dict[ImagePart, List[ImageReference]] = {}
        image_part_order: List[ImagePart] = []

        for story_part in self._iter_story_parts():
            occurrence_index = 0
            for blip in self._iter_inline_picture_blips(story_part):
                image_part = self._embedded_image_part(story_part, blip)
                if image_part is None:
                    # -- a picture linked only externally is not an embedded image --
                    continue
                occurrence_index += 1
                if image_part not in references_by_image_part:
                    references_by_image_part[image_part] = []
                    image_part_order.append(image_part)
                references_by_image_part[image_part].append(
                    ImageReference(str(story_part.partname), occurrence_index)
                )

        return tuple(
            ImageReferenceGroup(image_part, tuple(references_by_image_part[image_part]))
            for image_part in image_part_order
        )

    def _embedded_image_part(
        self, story_part: StoryPart, blip: CT_Blip
    ) -> ImagePart | None:
        """Return the |ImagePart| embedded by `blip`, or |None| when `blip` has no
        embedded image (an externally-linked-only picture).

        Raise |ImageReferenceError| when `blip` claims an embedded image via
        ``r:embed`` but that relationship is missing or does not target an image
        part.
        """
        rId = blip.embed
        if rId is None:
            return None

        rels = story_part.rels
        try:
            rel = rels[rId]
        except KeyError:
            raise ImageReferenceError(
                "a:blip r:embed='%s' in part '%s' references a relationship that is "
                "missing from that part's relationship collection"
                % (rId, story_part.partname)
            )

        if rel.is_external:
            raise ImageReferenceError(
                "a:blip r:embed='%s' in part '%s' references an external "
                "relationship; r:embed must target an embedded image part "
                "(externally-linked images use r:link)"
                % (rId, story_part.partname)
            )

        if rel.reltype != RT.IMAGE:
            raise ImageReferenceError(
                "relationship '%s' referenced by a:blip r:embed in part '%s' has "
                "relationship type '%s', expected an image relationship ('%s')"
                % (rId, story_part.partname, rel.reltype, RT.IMAGE)
            )

        return cast(ImagePart, rel.target_part)

    def _iter_inline_picture_blips(self, story_part: StoryPart) -> List[CT_Blip]:
        """The ``a:blip`` elements of inline pictures in `story_part`, in document
        order, including pictures in tables and nested tables."""
        return cast("List[CT_Blip]", story_part.element.xpath(_INLINE_PICTURE_BLIP_XPATH))

    def _iter_story_parts(self) -> Iterator[StoryPart]:
        """Yield the document part followed by each distinct existing header/footer
        part, each exactly once.

        Header and footer parts are reached through the document part's
        relationships, so a header linked by multiple sections is traversed only its
        single actual part. No part is created for a section header/footer that is
        linked-to-previous.
        """
        document_part = self._document_part
        yield document_part

        seen_parts = {document_part}
        for rel in document_part.rels.values():
            if rel.is_external or rel.reltype not in (RT.HEADER, RT.FOOTER):
                continue
            story_part = rel.target_part
            if story_part in seen_parts:
                continue
            seen_parts.add(story_part)
            yield cast("StoryPart", story_part)
