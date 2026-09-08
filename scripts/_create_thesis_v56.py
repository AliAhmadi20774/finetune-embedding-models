"""Create thesis draft v56 with the completed train/test leakage audit."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
import xml.etree.ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "docs" / "porseman_thesis_ali_ahmadi_draft_v55.docx"
OUTPUT = ROOT / "docs" / "porseman_thesis_ali_ahmadi_draft_v56.docx"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}

ET.register_namespace("w", W)


def tag(name: str) -> str:
    return f"{{{W}}}{name}"


def text_of(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def replace_text(paragraph: ET.Element, text: str) -> None:
    properties = paragraph.find("w:pPr", NS)
    for child in list(paragraph):
        if child is not properties:
            paragraph.remove(child)
    run = ET.SubElement(paragraph, tag("r"))
    run_properties = ET.SubElement(run, tag("rPr"))
    ET.SubElement(run_properties, tag("rtl")).set(tag("val"), "1")
    value = ET.SubElement(run, tag("t"))
    value.text = text


def find_paragraph(paragraphs: list[ET.Element], prefix: str) -> ET.Element:
    for paragraph in paragraphs:
        if text_of(paragraph).startswith(prefix):
            return paragraph
    raise ValueError(f"Paragraph not found: {prefix}")


def main() -> None:
    with ZipFile(SOURCE) as source_zip:
        document_xml = source_zip.read("word/document.xml")
        root = ET.fromstring(document_xml)
        body = root.find("w:body", NS)
        if body is None:
            raise RuntimeError("Word document body was not found.")
        paragraphs = body.findall("w:p", NS)
        normal_template = next(p for p in paragraphs if len(text_of(p)) > 300)

        audit_heading = find_paragraph(paragraphs, "۳-۴-۵. تثبیت نمونه آزمون")
        replace_text(audit_heading, "۳-۴-۵. ممیزی هم‌پوشانی معنایی و حذف نشت داده")

        for old, new in (
            ("۳-۴-۶. داده آموزش و نقش آن در ریزتنظیم", "۳-۴-۷. داده آموزش و نقش آن در ریزتنظیم"),
            ("۳-۴-۷. محدودیت‌های تفکیک تصادفی", "۳-۴-۸. محدودیت‌های تفکیک تصادفی"),
            ("۳-۴-۸. جمع‌بندی", "۳-۴-۹. جمع‌بندی"),
        ):
            replace_text(find_paragraph(paragraphs, old), new)

        audit_texts = (
            "پس از تفکیک اولیه داده‌های آموزش و آزمون، برای کنترل نشت ناشی از پرسش‌های تکراری یا بازنویسی‌شده، ممیزی هم‌پوشانی معنایی انجام شد. برای این منظور، بردار dense همه پرسش‌های آموزش و آزمون با مدل پایه BGE-M3 ساخته و پس از نرمال‌سازی L2، شباهت کسینوسی میان هر پرسش آزمون و پرسش‌های آموزش محاسبه شد. از آن‌جا که بردارها نرمال‌سازی شده بودند، ضرب داخلی آن‌ها معادل شباهت کسینوسی بود.",
            "آستانه ۰٫۹ برای شناسایی موارد مشکوک در نظر گرفته شد. این بررسی ۱۴۸ زوج پرسش آموزش–آزمون با شباهت حداقل ۰٫۹ را آشکار کرد که مربوط به ۱۳۱ پرسش آزمون، معادل ۸٫۱۸ درصد کل مجموعه آزمون، بود. همچنین ۳۱ پرسش پس از نرمال‌سازی نگارشی فارسی، تطابق متنی دقیق با یک پرسش آموزشی داشتند. برای جلوگیری از برآورد خوش‌بینانه عملکرد مدل، ۱۳۱ پرسش آزمونِ دارای هم‌پوشانی حذف شدند. مجموعه آزمون نهایی در فایل porseman_test.csv شامل ۱٬۴۷۰ پرسش است؛ کد و گزارش ممیزی نیز برای بازتولیدپذیری نگهداری شده‌اند.",
        )
        insertion_index = list(body).index(audit_heading) + 1
        for audit_text in audit_texts:
            paragraph = deepcopy(normal_template)
            replace_text(paragraph, audit_text)
            body.insert(insertion_index, paragraph)
            insertion_index += 1
        stable_heading = deepcopy(audit_heading)
        replace_text(stable_heading, "۳-۴-۶. تثبیت نمونه آزمون")
        body.insert(insertion_index, stable_heading)

        limitation_heading = find_paragraph(body.findall("w:p", NS), "۴-۴-۱. خطر هم‌پوشانی میان آموزش و آزمون")
        limitation_text = deepcopy(normal_template)
        replace_text(
            limitation_text,
            "ممیزی پس از تفکیک اولیه نشان داد که بخشی از پرسش‌های آزمون، تکرار دقیق یا بازنویسی معنایی پرسش‌های آموزش هستند. این هم‌پوشانی می‌تواند معیارهای بازیابی را به‌طور خوش‌بینانه افزایش دهد؛ زیرا مدل در آموزش با پرسشی هم‌معنا یا تقریباً یکسان روبه‌رو شده است. به همین دلیل، همه موارد دارای شباهت کسینوسی حداقل ۰٫۹ از آزمون حذف و مجموعه آزمون نهایی به ۱٬۴۷۰ پرسش محدود شد. نتایج کمی نهایی باید با همین مجموعه آزمون پالایش‌شده گزارش شوند.",
        )
        body.insert(list(body).index(limitation_heading) + 1, limitation_text)

        future_heading = find_paragraph(body.findall("w:p", NS), "۵-۳-۱.")
        replace_text(future_heading, "۵-۳-۱. تفکیک گروهی مبتنی بر خوشه‌های معنایی پیش از آموزش و آزمون")

        revised_summary = next(
            p for p in body.findall("w:p", NS)
            if text_of(p).startswith("نتایج فعلی شواهد مناسبی از اثر ریزتنظیم دامنه‌ای")
        )
        replace_text(
            revised_summary,
            "نتایج فعلی شواهد مناسبی از اثر ریزتنظیم دامنه‌ای بر بازیابی پاسخ‌های فارسی پرسمان فراهم می‌کنند؛ بااین‌حال، ممیزی هم‌پوشانی معنایی نشان داد که تفکیک تصادفی اولیه دارای نمونه‌های آلوده بوده است. پس از حذف ۱۳۱ پرسش آزمونِ دارای شباهت کسینوسی حداقل ۰٫۹ با داده آموزش، مجموعه آزمون نهایی شامل ۱٬۴۷۰ پرسش شد. بنابراین، اعداد نهایی معیارهای Recall@1، Recall@5 و MRR@10 باید بر مبنای همین مجموعه آزمون پالایش‌شده تفسیر و گزارش شوند.",
        )

        updated_document_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        with ZipFile(OUTPUT, "w", ZIP_DEFLATED) as output_zip:
            for item in source_zip.infolist():
                content = updated_document_xml if item.filename == "word/document.xml" else source_zip.read(item.filename)
                output_zip.writestr(item, content)

    print(f"Written: {OUTPUT}")


if __name__ == "__main__":
    main()
