"""Search PDF text and extract pages without overwriting same-named sources."""
import os
import re
import tempfile
from pathlib import Path

import pymupdf


def unique_output_path(folder, source):
    source = Path(source)
    suffix = f"__{source.parent.name}" if re.fullmatch(r"\d{4}-\d{2}", source.parent.name) else ""
    stem = source.stem + suffix
    candidate = Path(folder) / f"{stem}.pdf"
    number = 2
    while candidate.exists():
        candidate = Path(folder) / f"{stem}__{number}.pdf"
        number += 1
    return candidate


def extract_pdf(source_path, output_folder, keyword, keep_contractor_page, context):
    keyword = keyword.strip()
    if not keyword:
        raise ValueError("請輸入搜尋關鍵字")
    pattern = re.compile(r"\s*".join(re.escape(char) for char in keyword), re.IGNORECASE)
    context.check_cancelled()
    with pymupdf.open(source_path) as source:
        if source.needs_pass:
            raise ValueError("PDF 已加密，請先解密後再處理")
        matches, contractor = [], []
        for number, page in enumerate(source):
            context.check_cancelled()
            text = page.get_text()
            if pattern.search(text):
                matches.append(number)
            if keep_contractor_page and "承商駐院工程師" in re.sub(r"\s+", "", text):
                contractor.append(number)
            context.progress((number + 1) / max(1, source.page_count) * 0.7)
        if not matches:
            context.log(f"{Path(source_path).name} 未找到符合頁面，跳過")
            return "skipped"
        pages = sorted(set(matches + contractor))
        folder = Path(output_folder)
        folder.mkdir(parents=True, exist_ok=True)
        output = unique_output_path(folder, source_path)
        fd, temp = tempfile.mkstemp(prefix="extract-", suffix=".pdf", dir=folder)
        os.close(fd)
        try:
            with pymupdf.open() as extracted:
                for index, number in enumerate(pages):
                    context.check_cancelled()
                    extracted.insert_pdf(source, from_page=number, to_page=number)
                    context.progress(0.7 + (index + 1) / len(pages) * 0.29)
                extracted.save(temp)
            context.check_cancelled()
            os.replace(temp, output)
        except PermissionError as error:
            raise PermissionError("無法儲存 PDF，請關閉正在閱讀的檔案，並確認輸出資料夾可寫入") from error
        finally:
            Path(temp).unlink(missing_ok=True)
    context.log(f"已擷取 {len(pages)} 頁：{output.name}")
    return "extracted"
