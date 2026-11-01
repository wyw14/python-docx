"""Absolutize relative external hyperlink targets across a WordprocessingML package.

An external hyperlink is stored in a package part as a relationship of the
``.../hyperlink`` relationship type with ``TargetMode="External"``; the hyperlink
elements in the part XML reference that relationship by its Id (rId). Its target is a
URI reference, and when that reference is relative it is interpreted relative to the
on-disk location of the package -- a location python-docx has no knowledge of.

The :func:`absolutize_hyperlinks` function rewrites those relative references to
absolute URIs using an explicit base URI supplied by the caller, so links keep
working after the document is moved to another directory. Only relationship targets
are rewritten; the referencing XML (and therefore display text, run formatting and
relationship Ids) is never touched, links that already have absolute targets,
internal bookmark links (which use ``w:anchor`` rather than a relationship), and
field-code (``HYPERLINK``) links are all left as they were.

Resolution follows the URI reference rules of :rfc:`3986` via the standard-library
:mod:`urllib.parse` machinery, so query strings, fragments and percent-encoded
octets are all preserved. Neither the network nor the referenced file is accessed.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING
from urllib.parse import SplitResult, urljoin, urlsplit

from docx.opc.constants import RELATIONSHIP_TYPE as RT

if TYPE_CHECKING:
    from docx.opc.package import OpcPackage
    from docx.opc.part import Part


#: URI schemes accepted as the absolute base URI in this first version.
SUPPORTED_BASE_SCHEMES = ("http", "https", "file")


@dataclass(frozen=True)
class HyperlinkChange:
    """A single relative-to-absolute external-hyperlink rewrite.

    ``part`` is the package part that owns the relationship (its ``.partname``
    identifies e.g. the main document, a header or a footer). ``rId`` is the
    relationship Id, retained unchanged by the rewrite. ``original`` is the URI
    reference stored in the relationship before rewriting and ``absolute`` is the
    absolute URI produced by resolving it against the supplied base URI.
    """

    part: Part
    rId: str
    original: str
    absolute: str


@dataclass(frozen=True)
class HyperlinkSkip:
    """An external hyperlink relationship that could not be absolutized.

    ``part`` and ``rId`` identify the relationship, ``original`` is its target as
    stored, and ``reason`` describes why it could not be resolved against the base
    URI. The relationship is left untouched.
    """

    part: Part
    rId: str
    original: str
    reason: str


@dataclass(frozen=True)
class AbsolutizeHyperlinksResult:
    """The outcome of an absolutization run.

    ``changes`` contains one |HyperlinkChange| for each relative external hyperlink
    relationship that was rewritten (or, in a dry run, would be rewritten).
    ``skipped`` contains one |HyperlinkSkip| for each external hyperlink
    relationship whose target could not be resolved. Relationships that already had
    an absolute target are not listed in either sequence.
    """

    changes: tuple[HyperlinkChange, ...]
    skipped: tuple[HyperlinkSkip, ...]


def absolutize_hyperlinks(
    package: OpcPackage, base_uri: str, dry_run: bool = False
) -> AbsolutizeHyperlinksResult:
    """Resolve relative external hyperlink targets in `package` to absolute URIs.

    Every part in `package` (main document, headers, footers and all other parts)
    is inspected for an external relationship of the hyperlink relationship type.
    Each relative target is resolved against the absolute `base_uri` per
    :rfc:`3986`, preserving query parameters, percent-encoded octets and fragments.
    Targets that already carry a URI scheme (absolute web, file and other links),
    internal bookmark links and field-code hyperlinks (which have no relationship)
    are left unchanged.

    `base_uri` must be an absolute URI using one of the ``http``, ``https`` or
    ``file`` schemes. An invalid base raises |ValueError| (|TypeError| when it is not
    a string) *before* anything in the package is modified. Relative targets that
    cannot be resolved are recorded in the result's ``.skipped`` sequence and left
    untouched; they do not prevent other targets from being rewritten.

    When `dry_run` is |True| the package is not modified; ``.changes`` then reports
    the rewrites that a non-dry run would make. In both cases each relationship is
    considered at most once, even when multiple hyperlink elements share it, and
    relationship Ids are never changed.

    The network and the target files are never accessed and the rewrite survives
    saving and reopening the package.
    """
    # -- validate the base before touching anything, so an invalid base can never --
    # -- leave a partially rewritten package behind --
    base = _validate_base_uri(base_uri)

    changes: list[HyperlinkChange] = []
    skipped: list[HyperlinkSkip] = []

    # -- one relationship object per (part, rId), so a relationship shared by --
    # -- several hyperlink elements is naturally rewritten exactly once --
    for part in package.iter_parts():
        for rId, rel in part.rels.items():
            if not rel.is_external or rel.reltype != RT.HYPERLINK:
                continue

            original = rel.target_ref

            if not original.strip():
                skipped.append(HyperlinkSkip(part, rId, original, "target URI is empty"))
                continue

            try:
                # -- a URI reference that already carries a scheme is absolute --
                if urlsplit(original).scheme:
                    continue
                absolute = urljoin(base.geturl(), original)
                resolved_scheme = urlsplit(absolute).scheme
            except ValueError as e:
                skipped.append(HyperlinkSkip(part, rId, original, str(e)))
                continue

            if resolved_scheme != base.scheme:
                skipped.append(
                    HyperlinkSkip(
                        part,
                        rId,
                        original,
                        "target could not be resolved to an absolute URI",
                    )
                )
                continue

            changes.append(HyperlinkChange(part, rId, original, absolute))
            if not dry_run:
                rel.target_ref = absolute

    return AbsolutizeHyperlinksResult(tuple(changes), tuple(skipped))


def _validate_base_uri(base_uri: str) -> SplitResult:
    """Return `base_uri` parsed into components, raising if it is not a usable base.

    The base must be an absolute URI using an http, https or file scheme. http and
    https bases must name a host; file bases must have an absolute path (so e.g.
    Windows paths and relative references are rejected). Validation happens before
    any package content is modified.
    """
    if not isinstance(base_uri, str):
        raise TypeError("base URI must be a string, got '%s'" % type(base_uri).__name__)

    try:
        base = urlsplit(base_uri)
    except ValueError as e:
        raise ValueError("base URI %r is not a valid absolute URI (%s)" % (base_uri, e)) from e

    if not base.scheme:
        raise ValueError("base URI must be absolute, got relative reference %r" % base_uri)

    if base.scheme not in SUPPORTED_BASE_SCHEMES:
        raise ValueError(
            "base URI must use one of the %s schemes, got %r"
            % (", ".join(repr(s) for s in SUPPORTED_BASE_SCHEMES), base_uri)
        )

    if base.scheme in ("http", "https") and not base.netloc:
        raise ValueError("%s base URI must include a host name, got %r" % (base.scheme, base_uri))

    if base.scheme == "file" and not base.path.startswith("/"):
        raise ValueError("file base URI must have an absolute path, got %r" % base_uri)

    return base
