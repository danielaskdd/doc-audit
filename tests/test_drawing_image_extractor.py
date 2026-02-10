#!/usr/bin/env python3
"""
Unit tests for drawing_image_extractor.py.
"""

import sys
import zipfile
from pathlib import Path

from lxml import etree

# Add skills/doc-audit/scripts directory to path (must be before import)
_scripts_dir = Path(__file__).parent.parent / "skills" / "doc-audit" / "scripts"
sys.path.insert(0, str(_scripts_dir))

from drawing_image_extractor import (  # noqa: E402  # type: ignore
    create_drawing_context,
    extract_drawing_placeholder_from_element,
    normalize_drawing_placeholders_in_text,
)

NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _write_minimal_docx(path: Path):
    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="emf" ContentType="image/x-emf"/>
  <Default Extension="wmf" ContentType="image/x-wmf"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

    rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rIdEmf" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.emf"/>
  <Relationship Id="rIdExt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.com/img.wmf" TargetMode="External"/>
</Relationships>
"""

    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("word/document.xml", "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body/></w:document>")
        zf.writestr("word/_rels/document.xml.rels", rels)
        zf.writestr("word/media/image1.png", b"PNGDATA")
        zf.writestr("word/media/image2.emf", b"EMFDATA")


def _build_drawing_xml(docpr_id: str, docpr_name: str, rel_attr: str, rel_id: str, use_anchor: bool = False) -> etree._Element:
    container_tag = "wp:anchor" if use_anchor else "wp:inline"
    xml = f"""
<w:drawing xmlns:w="{NSMAP['w']}" xmlns:wp="{NSMAP['wp']}" xmlns:a="{NSMAP['a']}"
           xmlns:pic="{NSMAP['pic']}" xmlns:r="{NSMAP['r']}">
  <{container_tag}>
    <wp:docPr id="{docpr_id}" name="{docpr_name}"/>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:blipFill>
            <a:blip r:{rel_attr}="{rel_id}"/>
          </pic:blipFill>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </{container_tag}>
</w:drawing>
"""
    return etree.fromstring(xml)


def test_extract_embedded_inline_and_anchor_images(tmp_path: Path):
    docx_path = tmp_path / "sample.docx"
    output_path = tmp_path / "sample_blocks.jsonl"
    _write_minimal_docx(docx_path)

    ctx = create_drawing_context(str(docx_path), str(output_path))

    inline_drawing = _build_drawing_xml("1", "Inline Image", "embed", "rIdImg", use_anchor=False)
    inline_placeholder = extract_drawing_placeholder_from_element(
        inline_drawing, context=ctx, include_extended_attrs=True
    )

    assert 'id="1"' in inline_placeholder
    assert 'name="Inline Image"' in inline_placeholder
    assert 'path="sample_blocks.image/image1.png"' in inline_placeholder
    assert 'format="png"' in inline_placeholder
    assert (tmp_path / "sample_blocks.image" / "image1.png").read_bytes() == b"PNGDATA"

    anchor_drawing = _build_drawing_xml("2", "Anchor Emf", "embed", "rIdEmf", use_anchor=True)
    anchor_placeholder = extract_drawing_placeholder_from_element(
        anchor_drawing, context=ctx, include_extended_attrs=True
    )

    assert 'id="2"' in anchor_placeholder
    assert 'name="Anchor Emf"' in anchor_placeholder
    assert 'path="sample_blocks.image/image2.emf"' in anchor_placeholder
    assert 'format="emf"' in anchor_placeholder
    assert (tmp_path / "sample_blocks.image" / "image2.emf").read_bytes() == b"EMFDATA"


def test_extract_external_link_no_download_and_url_path(tmp_path: Path):
    docx_path = tmp_path / "sample.docx"
    output_path = tmp_path / "sample_blocks.jsonl"
    _write_minimal_docx(docx_path)

    ctx = create_drawing_context(str(docx_path), str(output_path))

    link_drawing = _build_drawing_xml("3", "External Wmf", "link", "rIdExt", use_anchor=False)
    placeholder = extract_drawing_placeholder_from_element(
        link_drawing, context=ctx, include_extended_attrs=True
    )

    assert 'id="3"' in placeholder
    assert 'name="External Wmf"' in placeholder
    assert 'path="https://example.com/img.wmf"' in placeholder
    assert 'format="wmf"' in placeholder

    image_dir = tmp_path / "sample_blocks.image"
    files = sorted(p.name for p in image_dir.iterdir())
    # External link is not downloaded, so no wmf file should exist locally.
    assert "img.wmf" not in files


def test_extract_prefers_link_when_blip_has_both_link_and_embed(tmp_path: Path):
    docx_path = tmp_path / "sample.docx"
    output_path = tmp_path / "sample_blocks.jsonl"
    _write_minimal_docx(docx_path)

    ctx = create_drawing_context(str(docx_path), str(output_path))

    both_xml = f"""
<w:drawing xmlns:w="{NSMAP['w']}" xmlns:wp="{NSMAP['wp']}" xmlns:a="{NSMAP['a']}"
           xmlns:pic="{NSMAP['pic']}" xmlns:r="{NSMAP['r']}">
  <wp:inline>
    <wp:docPr id="4" name="Linked With Cache"/>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:blipFill>
            <a:blip r:embed="rIdImg" r:link="rIdExt"/>
          </pic:blipFill>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>
"""
    drawing_elem = etree.fromstring(both_xml)
    placeholder = extract_drawing_placeholder_from_element(
        drawing_elem, context=ctx, include_extended_attrs=True
    )

    assert 'id="4"' in placeholder
    assert 'name="Linked With Cache"' in placeholder
    assert 'path="https://example.com/img.wmf"' in placeholder
    assert 'format="wmf"' in placeholder
    # Link must win, so embedded cache image should not be exported.
    assert not (tmp_path / "sample_blocks.image" / "image1.png").exists()


def test_drawing_without_blip_keeps_placeholder_only(tmp_path: Path):
    docx_path = tmp_path / "sample.docx"
    output_path = tmp_path / "sample_blocks.jsonl"
    _write_minimal_docx(docx_path)

    ctx = create_drawing_context(str(docx_path), str(output_path))

    drawing_xml = f"""
<w:drawing xmlns:w="{NSMAP['w']}" xmlns:wp="{NSMAP['wp']}" xmlns:a="{NSMAP['a']}" xmlns:r="{NSMAP['r']}">
  <wp:inline>
    <wp:docPr id="9" name="Chart 1"/>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/>
    </a:graphic>
  </wp:inline>
</w:drawing>
"""
    drawing_elem = etree.fromstring(drawing_xml)
    placeholder = extract_drawing_placeholder_from_element(
        drawing_elem, context=ctx, include_extended_attrs=True
    )

    assert placeholder == '<drawing id="9" name="Chart 1" />'


def test_normalize_drawing_placeholders_for_matching():
    text = (
        'A <drawing id="8" name="Image 8" '
        'path="sample_blocks.image/image8.emf" format="emf" /> B'
    )
    normalized = normalize_drawing_placeholders_in_text(
        text,
        include_extended_attrs=False,
    )
    assert normalized == 'A <drawing id="8" name="Image 8" /> B'
