# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false

"""Integration and unit tests for absolutizing relative external hyperlinks."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from docx import Document
from docx.hyperlink import (
    SUPPORTED_BASE_SCHEMES,
    AbsolutizeHyperlinksResult,
    absolutize_hyperlinks,
)
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import Part
from docx.oxml.parser import parse_xml
from docx.text.paragraph import Paragraph

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_HLINK_FRAGS_DOCX = (
    Path(__file__).parents[1] / "features" / "steps" / "test_files" / "par-hlink-frags.docx"
)


def _hyperlink_element(rId: str, text: str, bold: bool = False):
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return parse_xml(
        f'<w:hyperlink xmlns:w="{_W}" xmlns:r="{_R}" r:id="{rId}">'
        f"<w:r>{rpr}<w:t>{text}</w:t></w:r>"
        f"</w:hyperlink>"
    )


def _bookmark_hyperlink_element(anchor: str, text: str):
    return parse_xml(
        f'<w:hyperlink xmlns:w="{_W}" w:anchor="{anchor}">'
        f"<w:r><w:t>{text}</w:t></w:r>"
        f"</w:hyperlink>"
    )


def _field_hyperlink_element(instr: str, text: str = "field link"):
    escaped_instr = instr.replace("&", "&amp;").replace('"', "&quot;")
    return parse_xml(
        f'<w:fldSimple xmlns:w="{_W}" w:instr="{escaped_instr}">'
        f"<w:r><w:t>{text}</w:t></w:r>"
        f"</w:fldSimple>"
    )


def _add_hyperlink(part: Part, paragraph: Paragraph, target: str, text: str, bold: bool = False):
    """Add an external hyperlink relationship referencing `target` to `paragraph`."""
    rId = part.relate_to(target, RT.HYPERLINK, is_external=True)
    paragraph._p.append(_hyperlink_element(rId, text, bold))
    return rId


class DescribeAbsolutizeHyperlinks:
    """Integration suite for `Document.absolutize_hyperlinks()`."""

    def it_absolutizes_relative_targets_in_the_document_body(self):
        document = Document()
        paragraph = document.add_paragraph()
        rId = _add_hyperlink(document.part, paragraph, "docs/page.html?q=1#sec", "body link")

        result = document.absolutize_hyperlinks("https://example.com/site/report.docx")

        assert isinstance(result, AbsolutizeHyperlinksResult)
        assert len(result.changes) == 1
        assert result.skipped == ()
        change = result.changes[0]
        assert change.part is document.part
        assert change.rId == rId
        assert change.original == "docs/page.html?q=1#sec"
        assert change.absolute == "https://example.com/site/docs/page.html?q=1#sec"
        # -- the rewrite is visible through the hyperlink proxy, rId retained --
        hyperlink = paragraph.hyperlinks[0]
        assert hyperlink.address == change.absolute
        assert hyperlink._hyperlink.rId == rId

    def it_preserves_query_parameters_fragments_and_percent_escaping(self):
        document = Document()
        paragraph = document.add_paragraph()
        target = "a%20b/c%3Fd.html?x=1%202#frag%21"
        _add_hyperlink(document.part, paragraph, target, "encoded")

        result = document.absolutize_hyperlinks("https://h.test/docs/report.docx")

        assert result.changes[0].original == target
        assert result.changes[0].absolute == "https://h.test/docs/a%20b/c%3Fd.html?x=1%202#frag%21"

    @pytest.mark.parametrize(
        ("scheme", "base_uri", "target", "expected"),
        [
            (
                "http",
                "http://example.com/docs/report.docx",
                "../img/a.png",
                "http://example.com/img/a.png",
            ),
            (
                "https",
                "https://example.com/site/sub/",
                "../img/a.png?raw=1#p1",
                "https://example.com/site/img/a.png?raw=1#p1",
            ),
            (
                "file",
                "file:///C:/docs/report.docx",
                "../img/a%20b.png?raw=1#p1",
                "file:///C:/img/a%20b.png?raw=1#p1",
            ),
            (
                "file",
                "file://server/share/docs/report.docx",
                "../img/a.png",
                "file://server/share/img/a.png",
            ),
        ],
    )
    def it_supports_http_https_and_file_bases(
        self, scheme: str, base_uri: str, target: str, expected: str
    ):
        assert scheme in SUPPORTED_BASE_SCHEMES
        document = Document()
        paragraph = document.add_paragraph()
        _add_hyperlink(document.part, paragraph, target, "link")

        result = document.absolutize_hyperlinks(base_uri)

        assert result.changes[0].absolute == expected
        assert paragraph.hyperlinks[0].address == expected

    def it_leaves_already_absolute_targets_unchanged(self):
        document = Document()
        targets = [
            "https://web.example/x",
            "http://web.example/y",
            "file:///C:/abs/file.docx",
            "mailto:someone@example.com",
        ]
        rIds = []
        for target in targets:
            rIds.append(_add_hyperlink(document.part, document.add_paragraph(), target, "link"))

        result = document.absolutize_hyperlinks("https://example.com/docs/report.docx")

        assert result.changes == ()
        assert result.skipped == ()
        for rId, target in zip(rIds, targets):
            assert document.part.rels[rId].target_ref == target

    def it_leaves_internal_bookmark_links_unchanged(self):
        document = Document()
        paragraph = document.add_paragraph()
        paragraph._p.append(_bookmark_hyperlink_element("_Toc12345", "jump to heading"))

        result = document.absolutize_hyperlinks("https://example.com/docs/report.docx")

        assert result.changes == ()
        assert result.skipped == ()
        hyperlink = paragraph.hyperlinks[0]
        assert hyperlink.address == ""
        assert hyperlink.fragment == "_Toc12345"
        assert hyperlink.url == ""
        assert hyperlink.text == "jump to heading"

    def it_leaves_field_code_hyperlinks_unchanged(self):
        document = Document()
        paragraph = document.add_paragraph()
        instr = ' HYPERLINK "rel/field.html" '
        paragraph._p.append(_field_hyperlink_element(instr))

        result = document.absolutize_hyperlinks("https://example.com/docs/report.docx")

        # -- a HYPERLINK field carries its target in instrText, not a relationship --
        assert result.changes == ()
        assert result.skipped == ()
        assert document.part.element.xpath("//w:fldSimple/@w:instr") == [instr]

    def it_absolutizes_hyperlinks_in_header_and_footer_parts(self):
        document = Document()
        header = document.sections[0].header
        footer = document.sections[0].footer
        header_paragraph = header.paragraphs[0]
        footer_paragraph = footer.paragraphs[0]
        header_rId = _add_hyperlink(header.part, header_paragraph, "hdr/page.html", "header link")
        footer_rId = _add_hyperlink(footer.part, footer_paragraph, "ftr/page.html", "footer link")
        body_rId = _add_hyperlink(
            document.part, document.add_paragraph(), "body/page.html", "body link"
        )

        result = document.absolutize_hyperlinks("https://example.com/docs/report.docx")

        changed = {(c.part.partname, c.rId, c.original, c.absolute) for c in result.changes}
        assert changed == {
            (
                "/word/document.xml",
                body_rId,
                "body/page.html",
                "https://example.com/docs/body/page.html",
            ),
            (
                "/word/header1.xml",
                header_rId,
                "hdr/page.html",
                "https://example.com/docs/hdr/page.html",
            ),
            (
                "/word/footer1.xml",
                footer_rId,
                "ftr/page.html",
                "https://example.com/docs/ftr/page.html",
            ),
        }
        assert header_paragraph.hyperlinks[0].address == ("https://example.com/docs/hdr/page.html")
        assert footer_paragraph.hyperlinks[0].address == ("https://example.com/docs/ftr/page.html")

    def it_rewrites_a_shared_relationship_only_once_without_losing_text_or_format(self):
        document = Document()
        paragraph_1 = document.add_paragraph()
        paragraph_2 = document.add_paragraph()
        # -- relate_to() dedupes on target, so both hyperlinks share one rId --
        shared_rId_1 = document.part.relate_to("shared/page.html", RT.HYPERLINK, is_external=True)
        shared_rId_2 = document.part.relate_to("shared/page.html", RT.HYPERLINK, is_external=True)
        assert shared_rId_1 == shared_rId_2
        paragraph_1._p.append(_hyperlink_element(shared_rId_1, "plain link"))
        paragraph_2._p.append(_hyperlink_element(shared_rId_1, "bold link", bold=True))

        result = document.absolutize_hyperlinks("https://example.com/docs/report.docx")

        assert len(result.changes) == 1
        change = result.changes[0]
        assert change.rId == shared_rId_1
        assert change.original == "shared/page.html"
        assert change.absolute == "https://example.com/docs/shared/page.html"
        # -- both hyperlink elements keep their rId, text and run formatting --
        h1, h2 = paragraph_1.hyperlinks[0], paragraph_2.hyperlinks[0]
        assert h1.address == h2.address == change.absolute
        assert h1._hyperlink.rId == h2._hyperlink.rId == shared_rId_1
        assert h1.text == "plain link"
        assert h2.text == "bold link"
        assert h1.runs[0].font.bold is None
        assert h2.runs[0].font.bold is True

    def it_does_not_modify_the_document_in_dry_run(self):
        document = Document()
        paragraph = document.add_paragraph()
        rId = _add_hyperlink(document.part, paragraph, "docs/page.html", "link")

        result = document.absolutize_hyperlinks(
            "https://example.com/docs/report.docx", dry_run=True
        )

        assert len(result.changes) == 1
        assert result.changes[0].absolute == "https://example.com/docs/docs/page.html"
        assert document.part.rels[rId].target_ref == "docs/page.html"
        assert paragraph.hyperlinks[0].address == "docs/page.html"

        # -- the reported plan is what a real run applies --
        applied = document.absolutize_hyperlinks("https://example.com/docs/report.docx")
        assert applied.changes[0].absolute == result.changes[0].absolute
        assert document.part.rels[rId].target_ref == "https://example.com/docs/docs/page.html"

    def it_lists_unresolvable_targets_as_skipped_and_keeps_working(self):
        document = Document()
        broken_paragraph = document.add_paragraph()
        empty_rId = document.part.relate_to("", RT.HYPERLINK, is_external=True)
        broken_paragraph._p.append(_hyperlink_element(empty_rId, "empty target"))
        good_paragraph = document.add_paragraph()
        good_rId = _add_hyperlink(document.part, good_paragraph, "docs/page.html", "good")

        result = document.absolutize_hyperlinks("https://example.com/docs/report.docx")

        assert len(result.skipped) == 1
        skip = result.skipped[0]
        assert skip.part is document.part
        assert skip.rId == empty_rId
        assert skip.original == ""
        assert skip.reason
        assert document.part.rels[empty_rId].target_ref == ""
        assert [c.rId for c in result.changes] == [good_rId]
        assert good_paragraph.hyperlinks[0].address == "https://example.com/docs/docs/page.html"

    def it_skips_relative_targets_that_cannot_be_parsed(self):
        document = Document()
        broken_paragraph = document.add_paragraph()
        broken_rId = _add_hyperlink(document.part, broken_paragraph, "//[bad/x", "broken host")
        good_paragraph = document.add_paragraph()
        good_rId = _add_hyperlink(document.part, good_paragraph, "docs/page.html", "good")

        result = document.absolutize_hyperlinks("https://example.com/docs/report.docx")

        assert [s.rId for s in result.skipped] == [broken_rId]
        assert result.skipped[0].reason  # -- an explanation is provided --
        assert document.part.rels[broken_rId].target_ref == "//[bad/x"
        assert [c.rId for c in result.changes] == [good_rId]
        assert good_paragraph.hyperlinks[0].address == ("https://example.com/docs/docs/page.html")

    def it_raises_on_an_invalid_base_before_modifying_anything(self):
        document = Document()
        paragraph = document.add_paragraph()
        rId = _add_hyperlink(document.part, paragraph, "docs/page.html", "link")

        for bad_base in ("rel/x", "ftp://h/x", "file:rel/x", "http:///x", r"C:\docs\x.docx"):
            with pytest.raises(ValueError, match="base URI"):
                document.absolutize_hyperlinks(bad_base)
            # -- a rejected run leaves no change behind --
            assert document.part.rels[rId].target_ref == "docs/page.html"

        for bad_base in (None, 42):
            with pytest.raises(TypeError):
                document.absolutize_hyperlinks(bad_base)

    def it_persists_rewritten_addresses_through_save_and_reopen(self, tmp_path):
        document = Document()
        paragraph = document.add_paragraph()
        rId = _add_hyperlink(document.part, paragraph, "docs/page.html?q=1#sec", "saved link")
        absolute = "https://example.com/docs/docs/page.html?q=1#sec"
        document.absolutize_hyperlinks("https://example.com/docs/report.docx")

        path = tmp_path / "rewritten.docx"
        document.save(path)

        reopened = Document(path)
        hyperlink = reopened.paragraphs[-1].hyperlinks[0]
        assert hyperlink.address == absolute
        assert hyperlink._hyperlink.rId == rId
        assert hyperlink.text == "saved link"
        # -- the rewritten target is actually serialized in the package --
        with zipfile.ZipFile(path) as z:
            rels_xml = z.read("word/_rels/document.xml.rels").decode("utf-8")
        assert f'Target="{absolute}"' in rels_xml
        assert 'Target="docs/page.html' not in rels_xml

    def it_only_rewrites_the_relative_link_of_a_word_produced_fixture(self):
        document = Document(str(_HLINK_FRAGS_DOCX))
        document_part = document.part

        result = document.absolutize_hyperlinks("file:///C:/docs/report.docx")

        # -- fixture contains four absolute web links and one relative file link, rId8 --
        assert [(c.rId, c.original, c.absolute) for c in result.changes] == [
            ("rId8", "court-exif.jpg", "file:///C:/docs/court-exif.jpg")
        ]
        assert result.skipped == ()
        assert document_part.rels["rId4"].target_ref == "https://foo.com"
        assert document_part.rels["rId8"].target_ref == "file:///C:/docs/court-exif.jpg"

    def it_accepts_a_package_directly(self):
        document = Document()
        paragraph = document.add_paragraph()
        rId = _add_hyperlink(document.part, paragraph, "docs/page.html", "link")

        result = absolutize_hyperlinks(
            document.part.package, "https://example.com/docs/report.docx"
        )

        assert len(result.changes) == 1
        assert result.changes[0].rId == rId
        assert result.changes[0].absolute == "https://example.com/docs/docs/page.html"
        assert document.part.rels[rId].target_ref == ("https://example.com/docs/docs/page.html")


class DescribeBaseUriValidation:
    """Unit coverage for the accepted and rejected base URI forms."""

    @pytest.mark.parametrize(
        "base_uri",
        [
            "http://example.com/report.docx",
            "https://example.com/docs/report.docx?x=1",
            "file:///C:/docs/report.docx",
            "file:///docs/report.docx",
            "file://server/share/report.docx",
        ],
    )
    def it_accepts_absolute_http_https_and_file_URIs(self, base_uri: str):
        document = Document()
        result = document.absolutize_hyperlinks(base_uri, dry_run=True)
        assert isinstance(result.changes, tuple)

    @pytest.mark.parametrize(
        "base_uri",
        [
            "docs/report.docx",
            "//example.com/docs/report.docx",
            "ftp://example.com/x",
            "mailto:a@b.com",
            "file:relative/path.docx",
            "http:///nohost",
            "https:///nohost",
            "C:\\docs\\report.docx",
            "http://[::1",
        ],
    )
    def it_rejects_relative_unsupported_or_malformed_bases(self, base_uri: str):
        document = Document()
        with pytest.raises(ValueError, match="base URI"):
            document.absolutize_hyperlinks(base_uri, dry_run=True)

    def it_rejects_non_string_bases(self):
        document = Document()
        with pytest.raises(TypeError):
            document.absolutize_hyperlinks(42)  # type: ignore[arg-type]
        with pytest.raises(TypeError):
            document.absolutize_hyperlinks(None)  # type: ignore[arg-type]
