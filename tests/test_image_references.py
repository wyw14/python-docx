# pyright: reportPrivateUsage=false

"""Integration and unit tests for the embedded-image reference inventory.

The behavior tests use real in-memory .docx packages built with python-docx (and, for
malformed or unsupported constructs, light-weight ZIP/XML surgery) so the inventory is
exercised through the same loading path used in production.
"""

from __future__ import annotations

import io
import re
import zipfile
from typing import Dict, List, Tuple

import pytest

from docx import Document
from docx.enum.section import WD_HEADER_FOOTER
from docx.exceptions import ImageReferenceError
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.reference import ImageReference, ImageReferenceGroup, ImageReferences
from docx.shared import Inches

from .unitutil.file import test_file

JPEG_PATH = test_file("python-icon.jpeg")
PNG_PATH = test_file("monty-truth.png")

Inventory = List[Tuple[str, str, Tuple[Tuple[str, int], ...]]]


# -- package fixtures and helpers --------------------------------------------------------------


def _inventory(document) -> Inventory:
    return [
        (
            group.partname,
            group.content_type,
            tuple((ref.partname, ref.index) for ref in group.references),
        )
        for group in document.image_references
    ]


def _round_trip(document) -> bytes:
    # -- exercise the query before saving to ensure querying has no effect --
    list(document.image_references)
    file = io.BytesIO()
    document.save(file)
    return file.getvalue()


def _open(blob: bytes):
    return Document(io.BytesIO(blob))


def _zip_items(blob: bytes) -> Dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(blob)) as z:
        return {name: z.read(name) for name in z.namelist()}


def _zip_bytes(items: Dict[str, bytes]) -> bytes:
    file = io.BytesIO()
    with zipfile.ZipFile(file, "w", zipfile.ZIP_DEFLATED) as z:
        for name, item_blob in items.items():
            z.writestr(name, item_blob)
    return file.getvalue()


def _document_xml(items: Dict[str, bytes]) -> str:
    return items["word/document.xml"].decode("utf-8")


def _document_rels_xml(items: Dict[str, bytes]) -> str:
    return items["word/_rels/document.xml.rels"].decode("utf-8")


def _append_paragraphs_before_body_end(items: Dict[str, bytes], paragraphs_xml: str) -> None:
    xml = _document_xml(items)
    items["word/document.xml"] = xml.replace(
        "</w:body>", paragraphs_xml + "</w:body>"
    ).encode("utf-8")


def _add_external_image_rel(items: Dict[str, bytes], rId: str, url: str) -> None:
    rels = _document_rels_xml(items)
    rel = (
        '<Relationship Id="%s" Type="%s" Target="%s" TargetMode="External"/>'
        % (rId, RT.IMAGE, url)
    )
    items["word/_rels/document.xml.rels"] = rels.replace(
        "</Relationships>", rel + "</Relationships>"
    ).encode("utf-8")


def _inline_paragraph(
    doc_pr_id: int,
    name: str,
    graphic_data_xml: str,
    wp_decl: str = "",
) -> str:
    return (
        '<w:p><w:r><w:drawing><wp:inline %s distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="100" cy="100"/><wp:docPr id="%d" name="%s"/>'
        "<a:graphic>%s</a:graphic></wp:inline></w:drawing></w:r></w:p>"
        % (wp_decl, doc_pr_id, name, graphic_data_xml)
    )


_WP_DECL = (
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
    'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
)


def _linked_only_picture_paragraph(rId: str) -> str:
    graphic_data = (
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic><pic:blipFill><a:blip r:link="%s"/></pic:blipFill>'
        '<pic:spPr/></pic:pic></a:graphicData>' % rId
    )
    return _inline_paragraph(101, "Linked Picture", graphic_data, _WP_DECL)


def _chart_paragraph() -> str:
    graphic_data = (
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        '<c:chart r:id="rIdChartDoesNotExist"/>'
        "</a:graphicData>"
    )
    return _inline_paragraph(102, "Chart", graphic_data, _WP_DECL)


def _smartart_paragraph() -> str:
    graphic_data = (
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
        "<dgm:relIds/>"
        "</a:graphicData>"
    )
    return _inline_paragraph(103, "SmartArt", graphic_data, _WP_DECL)


def _embedded_picture_paragraph(rId: str) -> str:
    graphic_data = (
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic><pic:blipFill><a:blip r:embed="%s"/></pic:blipFill>'
        '<pic:spPr/></pic:pic></a:graphicData>' % rId
    )
    return _inline_paragraph(200, "Duplicate", graphic_data, _WP_DECL)


def _embed_and_link_picture_paragraph(embed_rId: str, link_rId: str) -> str:
    graphic_data = (
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic><pic:blipFill>'
        '<a:blip r:embed="%s" r:link="%s"/>'
        "</pic:blipFill><pic:spPr/></pic:pic></a:graphicData>"
        % (embed_rId, link_rId)
    )
    return _inline_paragraph(201, "Cached-And-Linked", graphic_data, _WP_DECL)


def _anchor_picture_paragraph(rId: str) -> str:
    return (
        '<w:p><w:r><w:drawing>'
        '<wp:anchor %s distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        'relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:extent cx="100" cy="100"/><wp:docPr id="202" name="Floating"/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic><pic:blipFill><a:blip r:embed="%s"/></pic:blipFill>'
        "<pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:anchor>"
        "</w:drawing></w:r></w:p>" % (_WP_DECL, rId)
    )


def _vml_picture_paragraph(rId: str) -> str:
    vml_decl = (
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    )
    return (
        '<w:p><w:r><w:pict>'
        '<v:shape id="VML1" %s style="width:10pt;height:10pt">'
        '<v:imagedata r:id="%s" o:title="vml"/>'
        "</v:shape></w:pict></w:r></w:p>" % (vml_decl, rId)
    )


def _first_embed_rId(items: Dict[str, bytes]) -> str:
    match = re.search(r'r:embed="(rId\d+)"', _document_xml(items))
    assert match is not None
    return match.group(1)


def _package_with_pictures() -> bytes:
    """Body picture, table pictures (incl. nested), two sections, header and footer.

    The JPEG is used in body, table, nested table, and header (sharing one part); the
    PNG is used once in a table cell and once in the footer (sharing another).
    Section two has no header/footer definitions (linked-to-previous).
    """
    document = Document()
    document.add_picture(JPEG_PATH, width=Inches(1))

    table = document.add_table(rows=1, cols=2)
    table.cell(0, 0).paragraphs[0].add_run().add_picture(JPEG_PATH, width=Inches(0.5))
    table.cell(0, 1).paragraphs[0].add_run().add_picture(PNG_PATH, width=Inches(0.5))
    nested = table.cell(0, 1).add_table(rows=1, cols=1)
    nested.cell(0, 0).paragraphs[0].add_run().add_picture(JPEG_PATH, width=Inches(0.4))

    document.add_section()

    header = document.sections[0].header
    header.is_linked_to_previous = False
    header.paragraphs[0].add_run().add_picture(JPEG_PATH, width=Inches(0.3))

    footer = document.sections[0].footer
    footer.is_linked_to_previous = False
    footer.paragraphs[0].add_run().add_picture(PNG_PATH, width=Inches(0.3))

    return _round_trip(document)


# -- DescribeImageReferences -------------------------------------------------------------------


class DescribeImageReferences:
    def it_returns_an_empty_inventory_for_a_document_with_no_pictures(self):
        document = Document()

        image_references = document.image_references

        assert isinstance(image_references, ImageReferences)
        assert len(image_references) == 0
        assert list(image_references) == []

    def it_reflects_pictures_added_after_an_earlier_query(self):
        document = Document()
        assert len(document.image_references) == 0

        document.add_picture(JPEG_PATH, width=Inches(1))
        header = document.sections[0].header
        header.is_linked_to_previous = False
        header.paragraphs[0].add_run().add_picture(JPEG_PATH, width=Inches(0.3))

        groups = list(document.image_references)

        assert len(groups) == 1
        assert [(r.partname, r.index) for r in groups[0].references] == [
            ("/word/document.xml", 1),
            ("/word/header1.xml", 1),
        ]

    def it_groups_references_by_image_part_with_partname_content_type_and_count(self):
        document = _open(_package_with_pictures())

        groups = list(document.image_references)

        assert [g.partname for g in groups] == [
            "/word/media/image1.jpeg",
            "/word/media/image2.png",
        ]
        assert [g.content_type for g in groups] == ["image/jpeg", "image/png"]
        assert [g.count for g in groups] == [4, 2]
        assert [g.reference_count for g in groups] == [4, 2]
        assert [len(g) for g in groups] == [4, 2]

    def it_records_the_host_partname_and_per_part_occurrence_index_for_each_reference(self):
        document = _open(_package_with_pictures())

        assert _inventory(document) == [
            (
                "/word/media/image1.jpeg",
                "image/jpeg",
                (
                    ("/word/document.xml", 1),
                    ("/word/document.xml", 2),
                    # -- index 3 in document.xml belongs to image2.png --
                    ("/word/document.xml", 4),
                    ("/word/header1.xml", 1),
                ),
            ),
            (
                "/word/media/image2.png",
                "image/png",
                (
                    ("/word/document.xml", 3),
                    ("/word/footer1.xml", 1),
                ),
            ),
        ]

    def it_counts_pictures_in_tables_and_nested_tables(self):
        document = Document()
        table = document.add_table(rows=1, cols=1)
        table.cell(0, 0).paragraphs[0].add_run().add_picture(PNG_PATH)
        nested = table.cell(0, 0).add_table(rows=1, cols=1)
        nested.cell(0, 0).paragraphs[0].add_run().add_picture(PNG_PATH)
        document = _open(_round_trip(document))

        group = document.image_references[0]

        assert group.partname == "/word/media/image1.png"
        assert group.count == 2
        assert [r.index for r in group.references] == [1, 2]
        assert {r.partname for r in group.references} == {"/word/document.xml"}

    def it_groups_a_picture_shared_between_body_and_header_into_one_group(self):
        document = Document()
        document.add_picture(JPEG_PATH, width=Inches(1))
        header = document.sections[0].header
        header.is_linked_to_previous = False
        header.paragraphs[0].add_run().add_picture(JPEG_PATH, width=Inches(0.3))
        document = _open(_round_trip(document))

        groups = list(document.image_references)

        assert len(groups) == 1
        (group,) = groups
        assert group.partname == "/word/media/image1.jpeg"
        assert group.count == 2
        assert [(r.partname, r.index) for r in group.references] == [
            ("/word/document.xml", 1),
            ("/word/header1.xml", 1),
        ]

    def it_traverses_a_header_part_once_when_multiple_sections_reference_it(self):
        document = Document()
        document.add_picture(JPEG_PATH, width=Inches(1))
        document.add_section()
        header = document.sections[0].header
        header.is_linked_to_previous = False
        header.paragraphs[0].add_run().add_picture(JPEG_PATH, width=Inches(0.3))
        # -- section 1 explicitly references the SAME header part as section 0 --
        rId = document.part.relate_to(header.part, RT.HEADER)
        document.sections[1]._sectPr.add_headerReference(WD_HEADER_FOOTER.PRIMARY, rId)
        document = _open(_round_trip(document))

        group = document.image_references[0]

        assert [(r.partname, r.index) for r in group.references] == [
            ("/word/document.xml", 1),
            ("/word/header1.xml", 1),
        ]

    def it_does_not_traverse_an_extra_part_for_a_section_linked_to_previous(self):
        # -- _package_with_pictures has two sections, only the first defines hdr/ftr --
        document = _open(_package_with_pictures())
        host_partnames = {
            ref.partname for group in document.image_references for ref in group.references
        }

        assert host_partnames == {
            "/word/document.xml",
            "/word/header1.xml",
            "/word/footer1.xml",
        }

    def it_reports_distinct_parts_separately_even_when_their_bytes_are_identical(self):
        blob = _package_with_pictures()
        items = _zip_items(blob)
        items["word/media/image9.jpeg"] = items["word/media/image1.jpeg"]
        rels = _document_rels_xml(items)
        rels = rels.replace(
            "</Relationships>",
            '<Relationship Id="rIdDup9" Type="%s" Target="media/image9.jpeg"/>'
            "</Relationships>" % RT.IMAGE,
        )
        items["word/_rels/document.xml.rels"] = rels.encode("utf-8")
        _append_paragraphs_before_body_end(items, _embedded_picture_paragraph("rIdDup9"))
        document = _open(_zip_bytes(items))

        groups = list(document.image_references)

        assert [g.partname for g in groups] == [
            "/word/media/image1.jpeg",
            "/word/media/image2.png",
            "/word/media/image9.jpeg",
        ]
        assert groups[0].count == 4
        assert groups[2].count == 1
        assert groups[2].references[0].partname == "/word/document.xml"
        assert groups[2].references[0].index == 5

    def it_ignores_externally_linked_pictures_without_following_the_link(self):
        blob = _package_with_pictures()
        items = _zip_items(blob)
        _add_external_image_rel(items, "rIdExt999", "http://example.invalid/linked.png")
        _append_paragraphs_before_body_end(items, _linked_only_picture_paragraph("rIdExt999"))
        document = _open(_zip_bytes(items))

        assert _inventory(document) == _inventory(_open(blob))

    def it_ignores_chart_and_smartart_inline_shapes(self):
        blob = _package_with_pictures()
        items = _zip_items(blob)
        _append_paragraphs_before_body_end(
            items, _chart_paragraph() + _smartart_paragraph()
        )
        document = _open(_zip_bytes(items))

        assert _inventory(document) == _inventory(_open(blob))

    def it_ignores_floating_anchor_pictures(self):
        blob = _package_with_pictures()
        items = _zip_items(blob)
        rId = _first_embed_rId(items)
        _append_paragraphs_before_body_end(items, _anchor_picture_paragraph(rId))
        document = _open(_zip_bytes(items))

        assert _inventory(document) == _inventory(_open(blob))

    def it_ignores_legacy_vml_pictures(self):
        blob = _package_with_pictures()
        items = _zip_items(blob)
        rId = _first_embed_rId(items)
        _append_paragraphs_before_body_end(items, _vml_picture_paragraph(rId))
        document = _open(_zip_bytes(items))

        assert _inventory(document) == _inventory(_open(blob))

    def it_counts_the_embedded_copy_of_a_picture_that_also_has_an_external_link(self):
        blob = _package_with_pictures()
        items = _zip_items(blob)
        embed_rId = _first_embed_rId(items)
        _add_external_image_rel(items, "rIdExt999", "http://example.invalid/linked.png")
        _append_paragraphs_before_body_end(
            items, _embed_and_link_picture_paragraph(embed_rId, "rIdExt999")
        )
        document = _open(_zip_bytes(items))

        group = document.image_references[0]
        assert group.count == 5
        # -- the embedded cache is counted as a new body occurrence; the external
        # -- link is never followed and the header occurrence is unaffected --
        assert [(r.partname, r.index) for r in group.references] == [
            ("/word/document.xml", 1),
            ("/word/document.xml", 2),
            ("/word/document.xml", 4),
            ("/word/document.xml", 5),
            ("/word/header1.xml", 1),
        ]

    def it_raises_a_clear_error_when_an_embed_relationship_is_missing(self):
        blob = _package_with_pictures()
        items = _zip_items(blob)
        xml = _document_xml(items)
        rId = re.search(r'r:embed="(rId\d+)"', xml).group(1)
        items["word/document.xml"] = xml.replace(
            'r:embed="%s"' % rId, 'r:embed="rIdAbsent123"', 1
        ).encode("utf-8")
        document = _open(_zip_bytes(items))

        with pytest.raises(ImageReferenceError, match="rIdAbsent123"):
            list(document.image_references)

    def it_raises_a_clear_error_when_an_embed_relationship_targets_a_non_image_part(self):
        blob = _package_with_pictures()
        items = _zip_items(blob)
        xml = _document_xml(items)
        rId = re.search(r'r:embed="(rId\d+)"', xml).group(1)
        rels = _document_rels_xml(items)
        replacement = (
            '<Relationship Id="%s" Type="%s" Target="styles.xml"/>'
            % (rId, RT.STYLES)
        )
        items["word/_rels/document.xml.rels"] = re.sub(
            r'<Relationship[^>]*Id="%s"[^>]*/>' % rId, replacement, rels
        ).encode("utf-8")
        document = _open(_zip_bytes(items))

        with pytest.raises(ImageReferenceError, match=rId):
            list(document.image_references)

    def it_raises_a_clear_error_when_an_embed_relationship_is_external(self):
        blob = _package_with_pictures()
        items = _zip_items(blob)
        xml = _document_xml(items)
        rId = re.search(r'r:embed="(rId\d+)"', xml).group(1)
        rels = _document_rels_xml(items)
        replacement = (
            '<Relationship Id="%s" Type="%s" Target="http://example.invalid/x.jpg" '
            'TargetMode="External"/>' % (rId, RT.IMAGE)
        )
        items["word/_rels/document.xml.rels"] = re.sub(
            r'<Relationship[^>]*Id="%s"[^>]*/>' % rId, replacement, rels
        ).encode("utf-8")
        document = _open(_zip_bytes(items))

        with pytest.raises(ImageReferenceError, match="external"):
            list(document.image_references)

    def it_does_not_change_document_xml_relationships_parts_or_image_data(self):
        document = _open(_package_with_pictures())

        def snapshot():
            package = document.part.package
            state = {}
            for part in package.iter_parts():
                state[str(part.partname)] = (part.content_type, part.blob, part.rels.xml)
            state["/"] = package.rels.xml
            return state

        before = snapshot()
        # -- query repeatedly, including indexed access and occurrence iteration --
        assert len(document.image_references) == 2
        for group in document.image_references:
            assert group.count == len(list(group.references))
        after = snapshot()

        assert before == after

    def it_produces_a_stable_inventory_after_save_and_reload(self):
        live_document = Document()
        live_document.add_picture(JPEG_PATH, width=Inches(1))
        table = live_document.add_table(rows=1, cols=1)
        table.cell(0, 0).paragraphs[0].add_run().add_picture(JPEG_PATH)
        table.cell(0, 0).paragraphs[0].add_run().add_picture(PNG_PATH)
        expected = _inventory(live_document)

        reloaded_once = _open(_round_trip(live_document))
        assert _inventory(reloaded_once) == expected

        reloaded_twice = _open(_round_trip(reloaded_once))
        assert _inventory(reloaded_twice) == expected

    def it_supports_indexed_access_and_its_groups_support_iteration_and_length(self):
        document = _open(_package_with_pictures())
        image_references = document.image_references

        assert isinstance(image_references[0], ImageReferenceGroup)
        assert isinstance(image_references[0].references[0], ImageReference)
        assert list(image_references[0]) == list(image_references[0].references)
        with pytest.raises(IndexError):
            image_references[9]

    def it_exposes_value_object_fields(self):
        document = _open(_package_with_pictures())
        group = document.image_references[0]
        reference = group.references[0]

        assert group.partname == "/word/media/image1.jpeg"
        assert group.content_type == "image/jpeg"
        assert group.count == 4
        assert reference.partname == "/word/document.xml"
        assert reference.index == 1
