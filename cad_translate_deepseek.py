"""使用 DeepSeek 将 DWG 中的中文 Text/MText 翻译为英文并另存。

需要 Windows + AutoCAD/ZWCAD COM。默认不覆盖原图。
"""

from __future__ import annotations

import argparse
import csv
import logging
import os
from pathlib import Path

from dotenv import load_dotenv
from openai import OpenAI
from openpyxl import load_workbook
from pyautocad import Autocad

from cad_batch_replace import dwg_files, replace_text


def load_glossary(path: Path | None) -> dict[str, str]:
    if not path:
        return {}
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = book.active
        result = {}
        for source, target in sheet.iter_rows(min_col=1, max_col=2, values_only=True):
            if source not in (None, "") and target is not None:
                result[str(source)] = str(target)
        return result
    finally:
        book.close()


def translate_one(client: OpenAI, model: str, value: str, cache: dict[str, str], target_language: str = "English") -> str:
    cache_key = f"{target_language}\0{value}"
    if cache_key in cache:
        return cache[cache_key]
    response = client.chat.completions.create(
        model=model,
        temperature=0,
        messages=[
            {
                "role": "system",
                "content": (
                    f"You translate Chinese CAD engineering drawing text into concise professional {target_language}. "
                    "Preserve numbers, units, symbols, line breaks and identifiers. Return only the translation."
                ),
            },
            {"role": "user", "content": value},
        ],
    )
    translated = (response.choices[0].message.content or value).strip()
    cache[cache_key] = translated
    return translated


def run(source_dir: Path, output_dir: Path, glossary_path: Path | None, provider: str = "deepseek", target_language: str = "English", api_key: str | None = None, base_url: str | None = None, model: str | None = None, report_path: Path | None = None, progress=None) -> tuple[int, int, int]:
    load_dotenv()
    api_key = api_key or os.getenv("DEEPSEEK_API_KEY")
    if not api_key:
        raise RuntimeError("请在 .env 中设置 DEEPSEEK_API_KEY")
    client = OpenAI(api_key=api_key, base_url=base_url or ("https://api.openai.com/v1" if provider == "chatgpt" else "https://api.deepseek.com"))
    model = model or ("gpt-4o-mini" if provider == "chatgpt" else "deepseek-chat")
    glossary = load_glossary(glossary_path)
    cache: dict[str, str] = {}
    output_dir.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(filename=output_dir / "cad_translate.log", level=logging.INFO, encoding="utf-8")
    acad = Autocad(create_if_not_exists=True, visible=True)
    processed = changed = failed = 0
    report_file = report_path or (output_dir / "translation.csv")
    report_file.parent.mkdir(parents=True, exist_ok=True)
    with report_file.open("w", newline="", encoding="utf-8-sig") as fp:
        report = csv.writer(fp)
        report.writerow(["中文原文", "英文译文", "来源"])
        files = list(dwg_files(source_dir))
        for index, source in enumerate(files, 1):
            destination = output_dir / source.relative_to(source_dir)
            destination.parent.mkdir(parents=True, exist_ok=True)
            document = None
            try:
                document = acad.app.Documents.Open(str(source.resolve()))
                acad.doc = document
                file_changed = 0
                for entity in acad.iter_objects(["Text", "MText"]):
                    original = str(entity.TextString)
                    translated = glossary.get(original)
                    origin = "glossary"
                    if translated is None:
                        translated = translate_one(client, model, original, cache, target_language)
                        origin = provider
                    if translated != original:
                        entity.TextString = translated
                        file_changed += 1
                    report.writerow([original, translated, origin])
                document.SaveAs(str(destination.resolve()))
                processed += 1
                changed += file_changed
                logging.info("完成 %s，修改 %d 个文字对象", source, file_changed)
            except Exception:
                failed += 1
                logging.exception("处理失败 %s", source)
            finally:
                if document is not None:
                    try:
                        document.Close(False)
                    except Exception:
                        logging.exception("关闭失败 %s", source)
            if progress:
                progress(index, len(files), source.name)
    return processed, changed, failed


def main() -> None:
    parser = argparse.ArgumentParser(description="使用 DeepSeek 将 DWG 中文文字翻译为英文")
    parser.add_argument("source", type=Path, help="DWG 源文件夹")
    parser.add_argument("output", type=Path, help="英文 DWG 输出文件夹")
    parser.add_argument("--glossary", type=Path, help="可选 Excel 术语表：第一列中文，第二列英文")
    parser.add_argument("--provider", choices=["deepseek", "chatgpt"], default="deepseek")
    parser.add_argument("--language", default="English", help="目标语言")
    args = parser.parse_args()
    result = run(args.source, args.output, args.glossary, args.provider, args.language)
    print(f"完成：处理 {result[0]} 个文件，修改 {result[1]} 个文字对象，失败 {result[2]} 个。")


if __name__ == "__main__":
    main()
