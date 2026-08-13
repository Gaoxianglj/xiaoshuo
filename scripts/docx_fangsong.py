"""WPS 优先的 DOCX 仿宋字体设置与嵌入工具。

本模块做两件事：
1. 把 WordprocessingML 中所有正文运行与默认样式的四类字体字段统一为“仿宋”；
2. 从本机 WPS 安装目录读取允许编辑嵌入的 FangS-SC.ttf，并按 OOXML 规范写入 DOCX。

字体文件只在构建时读取，不复制为项目内的独立资源。
"""

from __future__ import annotations

import argparse
import hashlib
import os
import struct
import tempfile
import uuid
import zipfile
from pathlib import Path

from docx.oxml.ns import qn
from lxml import etree


FONT_NAME = "仿宋"
FONT_ALT_NAME = "仿宋-简"
FONT_PART = "word/fonts/FangS-SC.odttf"
FONT_REL_TARGET = "fonts/FangS-SC.odttf"
FONT_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.obfuscatedFont"
FONT_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
FONT_TABLE_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
)
FONT_TABLE_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
)

WPS_FONT_CANDIDATES = (
    Path("/Applications/wpsoffice.app/Contents/Resources/office6/fonts/FangS-SC.ttf"),
    Path("/Applications/WPS Office.app/Contents/Resources/office6/fonts/FangS-SC.ttf"),
)

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

W = f"{{{NS_W}}}"
R = f"{{{NS_R}}}"
REL = f"{{{NS_REL}}}"
CT = f"{{{NS_CT}}}"

etree.register_namespace("w", NS_W)
etree.register_namespace("r", NS_R)


def set_run_font(run, size=None, bold=None, italic=None):
    """供 python-docx 生成器调用：同时写入四类字体字段。"""
    run.font.name = FONT_NAME
    if size is not None:
        run.font.size = size
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    fonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    for field in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(f"w:{field}"), FONT_NAME)
    fonts.set(qn("w:hint"), "eastAsia")


def configure_style_font(style, size=None):
    """供 python-docx 生成器调用：统一样式字体。"""
    style.font.name = FONT_NAME
    if size is not None:
        style.font.size = size
    fonts = style._element.get_or_add_rPr().get_or_add_rFonts()
    for field in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(f"w:{field}"), FONT_NAME)
    fonts.set(qn("w:hint"), "eastAsia")


def find_wps_font(explicit_path: str | Path | None = None) -> Path:
    candidates = []
    if explicit_path:
        candidates.append(Path(explicit_path))
    env_path = os.environ.get("FANGSONG_FONT_PATH")
    if env_path:
        candidates.append(Path(env_path))
    candidates.extend(WPS_FONT_CANDIDATES)

    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve()
    searched = "\n".join(f"- {path}" for path in candidates)
    raise FileNotFoundError(
        "未找到 WPS 仿宋字体 FangS-SC.ttf。检查 WPS 是否安装，或设置 "
        f"FANGSONG_FONT_PATH。\n已检查：\n{searched}"
    )


def read_embedding_permission(font_data: bytes) -> int:
    """读取 TrueType/OpenType OS/2.fsType。0 表示无嵌入限制。"""
    if len(font_data) < 12:
        raise ValueError("字体文件过短，无法读取 SFNT 表。")
    table_count = struct.unpack(">H", font_data[4:6])[0]
    for index in range(table_count):
        start = 12 + index * 16
        record = font_data[start : start + 16]
        if len(record) != 16:
            break
        tag, _checksum, offset, length = struct.unpack(">4sIII", record)
        if tag == b"OS/2":
            if length < 10 or offset + 10 > len(font_data):
                raise ValueError("字体 OS/2 表损坏，无法确认嵌入许可。")
            return struct.unpack(">H", font_data[offset + 8 : offset + 10])[0]
    raise ValueError("字体缺少 OS/2 表，无法确认嵌入许可。")


def assert_font_can_embed(font_data: bytes) -> int:
    fs_type = read_embedding_permission(font_data)
    if fs_type & 0x0002:
        raise PermissionError("字体标记为 Restricted License Embedding，禁止嵌入。")
    if fs_type & 0x0200:
        raise PermissionError("字体仅允许位图嵌入，不适合本项目 DOCX。")
    return fs_type


def _font_key(font_data: bytes) -> str:
    # 由字体摘要稳定生成，保证同一字体重复构建时 DOCX 结构可复现。
    value = uuid.UUID(bytes=hashlib.sha256(font_data).digest()[:16])
    return "{" + str(value).upper() + "}"


def _obfuscate_font(font_data: bytes, font_key: str) -> bytes:
    if len(font_data) < 32:
        raise ValueError("字体文件过短，无法执行 OOXML 字体混淆。")
    key = bytes.fromhex(font_key.strip("{}").replace("-", ""))[::-1]
    result = bytearray(font_data)
    for index in range(32):
        result[index] ^= key[index % 16]
    return bytes(result)


def _parse_xml(data: bytes):
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
    return etree.fromstring(data, parser=parser)


def _serialize_xml(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _set_rfonts(fonts) -> None:
    for field in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(W + field, FONT_NAME)
    fonts.set(W + "hint", "eastAsia")
    for field in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "csTheme"):
        fonts.attrib.pop(W + field, None)


def _ensure_rpr_font(rpr) -> None:
    fonts = rpr.find(W + "rFonts")
    if fonts is None:
        fonts = etree.Element(W + "rFonts")
        rpr.insert(0, fonts)
    _set_rfonts(fonts)


def _normalize_text_part(data: bytes, *, is_styles: bool = False) -> bytes:
    root = _parse_xml(data)

    for fonts in root.iter(W + "rFonts"):
        _set_rfonts(fonts)
    for rpr in root.iter(W + "rPr"):
        _ensure_rpr_font(rpr)
    for run in root.iter(W + "r"):
        rpr = run.find(W + "rPr")
        if rpr is None:
            rpr = etree.Element(W + "rPr")
            run.insert(0, rpr)
        _ensure_rpr_font(rpr)

    if is_styles:
        defaults = root.find(W + "docDefaults")
        if defaults is None:
            defaults = etree.Element(W + "docDefaults")
            root.insert(0, defaults)
        rpr_default = defaults.find(W + "rPrDefault")
        if rpr_default is None:
            rpr_default = etree.SubElement(defaults, W + "rPrDefault")
        rpr = rpr_default.find(W + "rPr")
        if rpr is None:
            rpr = etree.SubElement(rpr_default, W + "rPr")
        _ensure_rpr_font(rpr)

    return _serialize_xml(root)


def _ensure_font_table(data: bytes, font_key: str, rel_id: str) -> bytes:
    root = _parse_xml(data)

    target = None
    for font in root.findall(W + "font"):
        if font.get(W + "name") == FONT_NAME:
            target = font
            break
    if target is None:
        target = etree.SubElement(root, W + "font")
        target.set(W + "name", FONT_NAME)

    alt_name = target.find(W + "altName")
    if alt_name is None:
        alt_name = etree.Element(W + "altName")
        target.insert(0, alt_name)
    alt_name.set(W + "val", FONT_ALT_NAME)

    for tag in ("embedRegular", "embedBold", "embedItalic", "embedBoldItalic"):
        for element in target.findall(W + tag):
            target.remove(element)
    embedded = etree.SubElement(target, W + "embedRegular")
    embedded.set(R + "id", rel_id)
    embedded.set(W + "fontKey", font_key)

    if not any(font.get(W + "name") == FONT_ALT_NAME for font in root.findall(W + "font")):
        alias = etree.SubElement(root, W + "font")
        alias.set(W + "name", FONT_ALT_NAME)
        charset = etree.SubElement(alias, W + "charset")
        charset.set(W + "val", "86")
        family = etree.SubElement(alias, W + "family")
        family.set(W + "val", "auto")

    return _serialize_xml(root)


def _next_rel_id(root) -> str:
    used = {item.get("Id") for item in root.findall(REL + "Relationship")}
    number = 1
    while f"rId{number}" in used:
        number += 1
    return f"rId{number}"


def _ensure_font_relationship(data: bytes | None) -> tuple[bytes, str]:
    if data:
        root = _parse_xml(data)
    else:
        root = etree.Element(REL + "Relationships", nsmap={None: NS_REL})

    for relationship in root.findall(REL + "Relationship"):
        if (
            relationship.get("Type") == FONT_REL_TYPE
            and relationship.get("Target") == FONT_REL_TARGET
        ):
            return _serialize_xml(root), relationship.get("Id")

    rel_id = _next_rel_id(root)
    relationship = etree.SubElement(root, REL + "Relationship")
    relationship.set("Id", rel_id)
    relationship.set("Type", FONT_REL_TYPE)
    relationship.set("Target", FONT_REL_TARGET)
    return _serialize_xml(root), rel_id


def _ensure_document_font_table_relationship(data: bytes) -> bytes:
    root = _parse_xml(data)
    for relationship in root.findall(REL + "Relationship"):
        if relationship.get("Type") == FONT_TABLE_REL_TYPE:
            relationship.set("Target", "fontTable.xml")
            return _serialize_xml(root)
    relationship = etree.SubElement(root, REL + "Relationship")
    relationship.set("Id", _next_rel_id(root))
    relationship.set("Type", FONT_TABLE_REL_TYPE)
    relationship.set("Target", "fontTable.xml")
    return _serialize_xml(root)


def _ensure_content_types(data: bytes) -> bytes:
    root = _parse_xml(data)
    odttf = next(
        (item for item in root.findall(CT + "Default") if item.get("Extension") == "odttf"),
        None,
    )
    if odttf is None:
        default = etree.SubElement(root, CT + "Default")
        default.set("Extension", "odttf")
        default.set("ContentType", FONT_CONTENT_TYPE)
    else:
        odttf.set("ContentType", FONT_CONTENT_TYPE)
    if not any(
        item.get("PartName") == "/word/fontTable.xml" for item in root.findall(CT + "Override")
    ):
        override = etree.SubElement(root, CT + "Override")
        override.set("PartName", "/word/fontTable.xml")
        override.set("ContentType", FONT_TABLE_CONTENT_TYPE)
    return _serialize_xml(root)


def _ensure_settings(data: bytes) -> bytes:
    root = _parse_xml(data)
    if root.find(W + "embedTrueTypeFonts") is None:
        root.append(etree.Element(W + "embedTrueTypeFonts"))
    return _serialize_xml(root)


def _new_font_table() -> bytes:
    root = etree.Element(W + "fonts", nsmap={"w": NS_W, "r": NS_R})
    return _serialize_xml(root)


def _is_text_part(name: str) -> bool:
    if name in {"word/document.xml", "word/styles.xml", "word/numbering.xml"}:
        return True
    filename = Path(name).name
    return name.startswith("word/") and (
        filename.startswith("header")
        or filename.startswith("footer")
        or filename in {"footnotes.xml", "endnotes.xml", "comments.xml"}
    )


def embed_fangsong_docx(
    source: str | Path,
    output: str | Path | None = None,
    font_path: str | Path | None = None,
) -> Path:
    """统一字体并嵌入 WPS 仿宋。output 为空时原位安全替换。"""
    source = Path(source).resolve()
    destination = Path(output).resolve() if output else source
    font_file = find_wps_font(font_path)
    font_data = font_file.read_bytes()
    assert_font_can_embed(font_data)
    font_key = _font_key(font_data)
    obfuscated_font = _obfuscate_font(font_data, font_key)

    with zipfile.ZipFile(source, "r") as archive:
        entries = {item.filename: archive.read(item.filename) for item in archive.infolist()}

    required = {"[Content_Types].xml", "word/document.xml", "word/styles.xml", "word/settings.xml"}
    missing = sorted(required - entries.keys())
    if missing:
        raise ValueError(f"DOCX 缺少必要部件：{', '.join(missing)}")

    rels_name = "word/_rels/fontTable.xml.rels"
    entries[rels_name], rel_id = _ensure_font_relationship(entries.get(rels_name))
    entries["word/fontTable.xml"] = _ensure_font_table(
        entries.get("word/fontTable.xml", _new_font_table()), font_key, rel_id
    )
    entries["word/_rels/document.xml.rels"] = _ensure_document_font_table_relationship(
        entries["word/_rels/document.xml.rels"]
    )
    entries["[Content_Types].xml"] = _ensure_content_types(entries["[Content_Types].xml"])
    entries["word/settings.xml"] = _ensure_settings(entries["word/settings.xml"])
    entries[FONT_PART] = obfuscated_font

    for name in tuple(entries):
        if _is_text_part(name):
            entries[name] = _normalize_text_part(entries[name], is_styles=name == "word/styles.xml")

    destination.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary_name = tempfile.mkstemp(
        prefix=f".{destination.stem}.", suffix=".tmp.docx", dir=destination.parent
    )
    os.close(handle)
    temporary = Path(temporary_name)
    try:
        with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name, data in entries.items():
                archive.writestr(name, data)
        os.replace(temporary, destination)
    finally:
        if temporary.exists():
            temporary.unlink()
    return destination


def verify_fangsong_embedding(
    document: str | Path, font_path: str | Path | None = None
) -> dict[str, object]:
    """校验关系、字体键、字体内容及全部运行字体字段。"""
    document = Path(document).resolve()
    font_file = find_wps_font(font_path)
    original_font = font_file.read_bytes()

    with zipfile.ZipFile(document, "r") as archive:
        names = set(archive.namelist())
        required = {
            FONT_PART,
            "word/fontTable.xml",
            "word/_rels/fontTable.xml.rels",
            "word/settings.xml",
        }
        missing = sorted(required - names)
        if missing:
            raise AssertionError(f"缺少嵌入字体部件：{', '.join(missing)}")

        font_table = _parse_xml(archive.read("word/fontTable.xml"))
        target = next(
            (item for item in font_table.findall(W + "font") if item.get(W + "name") == FONT_NAME),
            None,
        )
        if target is None:
            raise AssertionError("fontTable.xml 中没有“仿宋”条目。")
        embedded = target.find(W + "embedRegular")
        if embedded is None:
            raise AssertionError("“仿宋”条目没有 embedRegular。")
        font_key = embedded.get(W + "fontKey")
        rel_id = embedded.get(R + "id")

        relationships = _parse_xml(archive.read("word/_rels/fontTable.xml.rels"))
        relation = next(
            (item for item in relationships.findall(REL + "Relationship") if item.get("Id") == rel_id),
            None,
        )
        if relation is None or relation.get("Target") != FONT_REL_TARGET:
            raise AssertionError("嵌入字体关系无效。")

        embedded_font = bytearray(archive.read(FONT_PART))
        restored = _obfuscate_font(bytes(embedded_font), font_key)
        if hashlib.sha256(restored).digest() != hashlib.sha256(original_font).digest():
            raise AssertionError("嵌入字体解混淆后与 WPS 字体不一致。")

        rfonts_count = 0
        for name in names:
            if not _is_text_part(name):
                continue
            root = _parse_xml(archive.read(name))
            for fonts in root.iter(W + "rFonts"):
                rfonts_count += 1
                for field in ("ascii", "hAnsi", "eastAsia", "cs"):
                    if fonts.get(W + field) != FONT_NAME:
                        raise AssertionError(f"{name} 中仍有非仿宋字体字段。")

        settings = _parse_xml(archive.read("word/settings.xml"))
        if settings.find(W + "embedTrueTypeFonts") is None:
            raise AssertionError("settings.xml 未启用 embedTrueTypeFonts。")

    return {
        "document": str(document),
        "font_source": str(font_file),
        "font_sha256": hashlib.sha256(original_font).hexdigest(),
        "font_key": font_key,
        "embedding_permission": f"0x{read_embedding_permission(original_font):04X}",
        "rfonts_checked": rfonts_count,
        "embedded_font_bytes": len(embedded_font),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="统一 DOCX 为仿宋并嵌入 WPS 字体。")
    parser.add_argument("document", type=Path, help="输入 DOCX")
    parser.add_argument("--output", type=Path, help="输出 DOCX；省略时原位更新")
    parser.add_argument("--font-path", type=Path, help="显式指定 FangS-SC.ttf")
    parser.add_argument("--verify-only", action="store_true", help="仅校验，不修改")
    args = parser.parse_args()

    if args.verify_only:
        target = args.output or args.document
    else:
        target = embed_fangsong_docx(args.document, args.output, args.font_path)
    result = verify_fangsong_embedding(target, args.font_path)
    for key, value in result.items():
        print(f"{key}: {value}")


if __name__ == "__main__":
    main()
