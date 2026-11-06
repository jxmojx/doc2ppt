"""OOXML 安全防护：拒绝含 DOCTYPE/ENTITY 声明的包（防 XXE）。

文档与模板均为用户输入。python-pptx / python-docx 内部用 lxml 解析 OOXML，
虽然默认不联网，但 DOCTYPE 内部实体可用于实体膨胀（billion laughs）或
外部实体引用。这里在解析前做一次包级扫描，命中即拒绝。
"""
import zipfile
from pathlib import Path

DANGEROUS = (b"<!DOCTYPE", b"<!ENTITY")


def guard_ooxml(path):
    """扫描 .docx/.pptx 包内所有 XML 部件，含 DOCTYPE/ENTITY 声明则抛 ValueError。"""
    p = Path(path)
    if p.suffix.lower() not in (".docx", ".pptx"):
        return
    try:
        zf = zipfile.ZipFile(p)
    except zipfile.BadZipFile as e:
        raise ValueError(f"不是有效的 OOXML 包: {p.name}") from e
    with zf:
        for name in zf.namelist():
            if not name.endswith((".xml", ".rels")):
                continue
            try:
                data = zf.read(name)
            except (KeyError, RuntimeError):
                continue
            if any(tag in data for tag in DANGEROUS):
                raise ValueError(
                    f"拒绝解析: {p.name} 的 {name} 含 DOCTYPE/ENTITY 声明（防 XXE 实体注入）"
                )
