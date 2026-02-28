# DOCX → HTML Web App

Небольшое веб-приложение на Python, которое принимает `.docx` файл и конвертирует его в HTML.

## Запуск

1. Откройте терминал в папке проекта (там, где лежит `app.py`).
2. Запустите сервер одной из команд:

```bash
python app.py
```

или (Linux/macOS):

```bash
python3 app.py
```

После запуска откройте в браузере: `http://localhost:8000`.

## Что поддерживается

- Абзацы
- Заголовки `Heading1..Heading6`
- Форматирование текста: **bold**, *italic*, underline
- Переносы строки
- Базовые таблицы (строки/ячейки, включая вложенные абзацы)

> Сложная верстка Word (продвинутые стили, колонки, фигуры и т.п.) в этой версии может конвертироваться ограниченно.

## Важно для Windows

Если вы видите ошибку вида `can't open file 'C:\Users\<user>\app.py'`, значит команда выполнялась не из папки проекта.
Сначала перейдите в директорию проекта, например:

```powershell
cd C:\path\to\WordHtml
python app.py
```

## Проверка конвертера из консоли

Команда вида `python3 - <<'PY' ...` работает в **bash**, но не в **PowerShell**.
Поэтому в PowerShell используйте here-string:

```powershell
@'
import io, zipfile
from app import docx_to_html

xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Before table</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>'''

buf = io.BytesIO()
with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('word/document.xml', xml)

result = docx_to_html(buf.getvalue())
print(result)
print('OK')
'@ | python -
```
