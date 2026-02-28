#!/usr/bin/env python3
from __future__ import annotations

import html
import io
import re
import sys
import zipfile
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Optional
from urllib.parse import parse_qs
import xml.etree.ElementTree as ET

HOST = "0.0.0.0"
PORT = 8000

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def style_to_tag(style: Optional[str]) -> str:
    if not style:
        return "p"
    m = re.match(r"Heading([1-6])", style)
    if m:
        return f"h{m.group(1)}"
    return "p"


def run_to_html(run: ET.Element) -> str:
    text_parts: list[str] = []
    for child in run:
        if child.tag == f"{{{W_NS}}}t":
            text_parts.append(html.escape(child.text or ""))
        elif child.tag == f"{{{W_NS}}}tab":
            text_parts.append("&emsp;")
        elif child.tag == f"{{{W_NS}}}br":
            text_parts.append("<br>")
    content = "".join(text_parts)

    rpr = run.find("w:rPr", NS)
    if rpr is not None:
        if rpr.find("w:b", NS) is not None:
            content = f"<strong>{content}</strong>"
        if rpr.find("w:i", NS) is not None:
            content = f"<em>{content}</em>"
        if rpr.find("w:u", NS) is not None:
            content = f"<u>{content}</u>"
    return content


def paragraph_to_html(paragraph: ET.Element) -> str:
    p_style = paragraph.find("w:pPr/w:pStyle", NS)
    style_val = p_style.attrib.get(f"{{{W_NS}}}val") if p_style is not None else None
    tag = style_to_tag(style_val)

    fragments: list[str] = []
    for run in paragraph.findall("w:r", NS):
        fragments.append(run_to_html(run))

    for link in paragraph.findall("w:hyperlink", NS):
        link_text = "".join(run_to_html(r) for r in link.findall("w:r", NS))
        fragments.append(link_text)

    body = "".join(fragments).strip()
    if not body:
        return ""
    return f"<{tag}>{body}</{tag}>"


def blocks_to_html(parent: ET.Element) -> list[str]:
    blocks: list[str] = []

    for element in parent:
        if element.tag == f"{{{W_NS}}}p":
            p_html = paragraph_to_html(element)
            if p_html:
                blocks.append(p_html)
        elif element.tag == f"{{{W_NS}}}tbl":
            table_html = table_to_html(element)
            if table_html:
                blocks.append(table_html)
        elif element.tag in (
            f"{{{W_NS}}}sdt",
            f"{{{W_NS}}}sdtContent",
            f"{{{W_NS}}}customXml",
        ):
            blocks.extend(blocks_to_html(element))
        elif list(element):
            # Резервный путь для контейнеров Word (например, smartTag):
            # спускаемся глубже и собираем известные блоки.
            blocks.extend(blocks_to_html(element))

    return blocks


def table_to_html(table: ET.Element) -> str:
    rows_html: list[str] = []

    for row in table.findall("w:tr", NS):
        cells_html: list[str] = []

        for cell in row.findall("w:tc", NS):
            cell_parts = blocks_to_html(cell)
            cell_body = "".join(cell_parts).strip()
            if not cell_body:
                cell_body = "&nbsp;"
            cells_html.append(f"<td>{cell_body}</td>")

        if not cells_html:
            continue

        # Отфильтровываем полностью пустые строки, которые часто встречаются в служебной верстке Word.
        if all(cell == "<td>&nbsp;</td>" for cell in cells_html):
            continue

        rows_html.append(f"<tr>{''.join(cells_html)}</tr>")

    if not rows_html:
        return ""

    return (
        '<table border="1" cellspacing="0" cellpadding="6">'
        f"{''.join(rows_html)}"
        "</table>"
    )


def docx_to_html(docx_bytes: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as archive:
        if "word/document.xml" not in archive.namelist():
            raise ValueError("Файл не содержит word/document.xml. Это точно DOCX?")
        xml_content = archive.read("word/document.xml")

    root = ET.fromstring(xml_content)
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("В документе отсутствует содержимое.")

    html_blocks = blocks_to_html(body)
    converted = "\n".join(html_blocks)
    return (
        "<!doctype html>\n"
        "<html lang=\"ru\">\n"
        "<head><meta charset=\"utf-8\"><title>Converted DOCX</title></head>\n"
        "<body>\n"
        f"{converted}\n"
        "</body></html>"
    )


def parse_multipart(content_type: str, body: bytes) -> tuple[Optional[str], Optional[bytes]]:
    match = re.search(r"boundary=([^;]+)", content_type)
    if not match:
        return None, None

    boundary = match.group(1).strip().strip('"').encode("utf-8")
    delimiter = b"--" + boundary
    parts = body.split(delimiter)

    for part in parts:
        part = part.strip()
        if not part or part == b"--":
            continue

        header_blob, sep, file_data = part.partition(b"\r\n\r\n")
        if not sep:
            continue

        headers = header_blob.decode("utf-8", errors="ignore")
        disp_match = re.search(r'filename="([^"]*)"', headers)
        name_match = re.search(r'name="([^"]*)"', headers)
        if not name_match or name_match.group(1) != "docx":
            continue

        filename = disp_match.group(1) if disp_match else ""
        file_data = file_data.rstrip(b"\r\n")
        return filename, file_data

    return None, None


PAGE_TEMPLATE = """<!doctype html>
<html lang=\"ru\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>DOCX → HTML</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 2rem; background: #f7f7fb; color: #222; }}
    .card {{ max-width: 900px; margin: 0 auto; background: #fff; border-radius: 12px; padding: 1.5rem; box-shadow: 0 2px 14px rgba(0,0,0,.08); }}
    h1 {{ margin-top: 0; }}
    .row {{ display: flex; gap: .6rem; flex-wrap: wrap; margin-bottom: 1rem; }}
    input[type=file] {{ padding: .4rem; border: 1px solid #ddd; border-radius: 8px; background: #fff; }}
    button {{ border: 0; border-radius: 8px; padding: .6rem 1rem; background: #2d6cdf; color: #fff; cursor: pointer; }}
    button:hover {{ background: #1d58c4; }}
    textarea {{ width: 100%; min-height: 260px; font-family: ui-monospace, SFMono-Regular, Menlo, monospace; border-radius: 8px; border: 1px solid #ddd; padding: .8rem; }}
    .error {{ color: #a11919; margin-bottom: .8rem; }}
    .hint {{ color: #4b5563; font-size: .95rem; margin-top: .6rem; }}
  </style>
</head>
<body>
  <div class=\"card\">
    <h1>Конвертер DOCX в HTML</h1>
    <p>Загрузите <code>.docx</code> файл и получите HTML-код.</p>
    {error}
    <form method=\"post\" enctype=\"multipart/form-data\">
      <div class=\"row\">
        <input type=\"file\" name=\"docx\" accept=\".docx\" required />
        <button type=\"submit\">Конвертировать</button>
      </div>
    </form>
    <p class=\"hint\">Если команда <code>python3 app.py</code> не работает в Windows, используйте <code>python app.py</code> из папки проекта.</p>
    {result}
  </div>
</body>
</html>
"""


class AppHandler(BaseHTTPRequestHandler):
    def _write_html(self, content: str, status: int = HTTPStatus.OK) -> None:
        body = content.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:  # noqa: N802
        self._write_html(PAGE_TEMPLATE.format(error="", result=""))

    def do_POST(self) -> None:  # noqa: N802
        content_type = self.headers.get("Content-Type", "")
        content_length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(content_length)

        error_html = ""
        result_html = ""

        filename: Optional[str] = None
        docx_bytes: Optional[bytes] = None

        if "multipart/form-data" in content_type:
            filename, docx_bytes = parse_multipart(content_type, body)
        elif "application/x-www-form-urlencoded" in content_type:
            data = parse_qs(body.decode("utf-8", errors="ignore"))
            if "docx" in data:
                filename = "upload.docx"
                docx_bytes = data["docx"][0].encode("utf-8")

        if not docx_bytes:
            error_html = '<p class="error">Файл не загружен или поврежден.</p>'
            self._write_html(PAGE_TEMPLATE.format(error=error_html, result=""), HTTPStatus.BAD_REQUEST)
            return

        filename = (filename or "").lower()
        if not filename.endswith(".docx"):
            error_html = '<p class="error">Нужен файл с расширением .docx.</p>'
            self._write_html(PAGE_TEMPLATE.format(error=error_html, result=""), HTTPStatus.BAD_REQUEST)
            return

        try:
            converted_html = docx_to_html(docx_bytes)
            escaped = html.escape(converted_html)
            result_html = (
                "<h2>Результат</h2>"
                f"<textarea readonly>{escaped}</textarea>"
                "<p>Скопируйте код из поля выше.</p>"
            )
        except Exception as exc:
            error_html = f'<p class="error">Ошибка конвертации: {html.escape(str(exc))}</p>'

        self._write_html(PAGE_TEMPLATE.format(error=error_html, result=result_html))


def main() -> None:
    server = ThreadingHTTPServer((HOST, PORT), AppHandler)
    print(f"Server started: http://{HOST}:{PORT}")
    server.serve_forever()


if __name__ == "__main__":
    if sys.version_info < (3, 10):
        raise SystemExit("Нужен Python 3.10 или новее")
    main()
