"""
DOCX → HTML 转换后端服务
使用 LibreOffice CLI 进行高质量转换，保留 Word 自动序号和排版。

启动方式:
    python server.py                  # 默认端口 8765
    python server.py 9000             # 自定义端口

依赖:
    - Python 3.8+
    - LibreOffice (soffice)
"""

import http.server
import json
import os
import re
import shutil
import subprocess
import sys
import time
import tempfile
import uuid
from urllib.parse import urlparse

# === LibreOffice detection ===
SOFFICE_PATHS = [
    r'C:\Program Files\LibreOffice\program\soffice.exe',
    r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
    '/usr/bin/soffice',
    '/usr/bin/libreoffice',
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
]

_soffice_path = None

def _kill_soffice():
    """Kill all soffice processes to prevent LibreOffice hangs."""
    import signal
    killed = False
    for proc_name in ('soffice.exe', 'soffice.bin'):
        try:
            result = subprocess.run(
                ['taskkill', '/F', '/IM', proc_name],
                capture_output=True, text=True, timeout=10
            )
            if result.returncode == 0:
                killed = True
        except Exception:
            pass
    if killed:
        time.sleep(1)  # let OS release file locks

def find_soffice():
    global _soffice_path
    if _soffice_path is not None:
        return _soffice_path or None
    # PATH
    found = shutil.which('soffice') or shutil.which('libreoffice')
    if found:
        _soffice_path = found
        return found
    for p in SOFFICE_PATHS:
        if os.path.isfile(p):
            _soffice_path = p
            return p
    _soffice_path = ''  # cache negative result
    return None

def convert_with_libreoffice(docx_path, output_dir):
    """Convert .docx to .html using LibreOffice CLI. Returns HTML body content."""
    soffice = find_soffice()
    if not soffice:
        return None, 'LibreOffice not found'

    # Kill any stuck soffice processes before starting (prevents hangs)
    _kill_soffice()

    cmd = [soffice, '--headless', '--norestore', '--convert-to', 'html', '--outdir', output_dir, docx_path]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

    if result.returncode != 0:
        return None, f'soffice exit code {result.returncode}: {result.stderr[:500]}'

    basename = os.path.splitext(os.path.basename(docx_path))[0] + '.html'
    output_path = os.path.join(output_dir, basename)

    if not os.path.isfile(output_path):
        return None, f'Output not found: {output_path}'

    with open(output_path, 'r', encoding='utf-8', errors='replace') as f:
        raw_html = f.read()

    # Extract <body> content
    body_match = re.search(r'<body[^>]*>([\s\S]*?)</body>', raw_html, re.IGNORECASE)
    body = body_match.group(1).strip() if body_match else raw_html

    # Clean LibreOffice HTML artifacts while preserving formatting
    # Remove XML declarations
    body = re.sub(r'<\?xml[^?]*\?>', '', body)
    # Remove namespace tags (e.g. <o:p>, <v:shape>)
    body = re.sub(r'</?\w+:\w+[^>]*>', '', body)
    # Remove conditional comments
    body = re.sub(r'<!--\[if[\s\S]*?<!\[endif\]-->', '', body)
    # Remove script tags
    body = re.sub(r'<script[^>]*>[\s\S]*?</script>', '', body, flags=re.IGNORECASE)
    # Remove link tags
    body = re.sub(r'<link[^>]*>', '', body, flags=re.IGNORECASE)
    # Remove VML shapes
    body = re.sub(r'<v:shape[\s\S]*?</v:shape>', '', body)
    body = re.sub(r'<v:imagedata[^>]*>', '', body)
    # Remove all class attributes (LO inserts "cjk", "Western", etc.)
    body = re.sub(r'\s+class="[^"]*"', '', body)
    # Remove lang attributes
    body = re.sub(r'\s+lang="[^"]*"', '', body)
    # Strip <font ...> and </font> tags, keep inner content
    body = re.sub(r'</?font[^>]*>', '', body, flags=re.IGNORECASE)
    # Clean <p> styles: remove LO artifact margin-right
    body = re.sub(r'margin-right:\s*-0\.08cm;\s*', '', body)
    body = re.sub(r';\s*margin-right:\s*-0\.08cm', '', body)
    # Normalize <p> tags to match reference format
    body = re.sub(r'<p\b[^>]*>', _normalize_p_tag, body)
    # Remove empty paragraphs at start/end
    body = re.sub(r'^\s*<p[^>]*>\s*</p>', '', body)
    body = re.sub(r'<p[^>]*>\s*</p>\s*$', '', body)

    # Fix LibreOffice nested lists
    body = fix_libreoffice_lists(body)

    # Collapse whitespace: convert tabs/newlines to single spaces
    body = re.sub(r'[\t\n\r]+', ' ', body)
    # Collapse multiple spaces
    body = re.sub(r'  +', ' ', body)

    return body, None


def _normalize_p_tag(match):
    """Normalize <p> tag attributes to match reference format."""
    attrs = match.group(0)
    # Extract align
    align_match = re.search(r'align="([^"]+)"', attrs)
    align = align_match.group(1) if align_match else 'justify'
    # Extract style (without margin-right which was already stripped)
    style_match = re.search(r'style="([^"]*)"', attrs)
    style = style_match.group(1).strip() if style_match else ''
    # Remove empty style remnants (trailing semicolons, etc.)
    style = re.sub(r';\s*;', ';', style)
    style = style.strip().strip(';').strip()
    if style:
        return f'<p align="{align}" style="{style}">'
    return f'<p align="{align}">'


def fix_libreoffice_lists(html):
    """Fix LibreOffice fragmented lists: each item as independent <ol><li> block.

    LO outputs each list item with depth via nesting levels:
      <ol><li>item</li></ol>                          depth=0  (I)
      <ol><ol type="a"><li>item</li></ol></ol>         depth=1  (A)
      <ol><ol type="a"><ol><li>item</li></ol></ol></ol> depth=2  (1)

    Strategy: strip wrapper <ol> levels, keep only innermost <ol><li>,
    use margin-left from the following translation <p> for indentation.
    """

    # Find all top-level <ol>...</ol> blocks
    segments = []  # ('text', html) or ('item', li_inner)
    pos = 0
    while pos < len(html):
        ol_start = html.find('<ol', pos)
        if ol_start < 0:
            segments.append(('text', html[pos:]))
            break
        if ol_start > pos:
            segments.append(('text', html[pos:ol_start]))

        d = 0; i = ol_start
        while i < len(html):
            m_open = re.match(r'<(?:ol|ul)[^>]*>', html[i:])
            m_close = re.match(r'</(?:ol|ul)>', html[i:])
            if m_open:
                d += 1; i += m_open.end()
            elif m_close:
                d -= 1; i += m_close.end()
                if d == 0: break
            else:
                i += 1
        block = html[ol_start:i]

        li_matches = re.findall(r'<li>([\s\S]*?)</li>', block)
        li_inner = li_matches[-1] if li_matches else ''
        segments.append(('item', li_inner))
        pos = i

    # Phase 2: Determine depth from margin-left of trailing <p>
    # margin-left: Ncm → depth = N-1 (1cm→0, 2cm→1, 3cm→2, ...)
    MARGIN_TO_DEPTH = {}
    CSS_LIST_STYLE = {0: 'upper-roman', 1: 'upper-alpha', 2: 'decimal', 3: 'lower-alpha', 4: 'lower-roman'}
    counters = {d: 0 for d in range(6)}
    result = []

    for idx, seg in enumerate(segments):
        if seg[0] == 'text':
            result.append(seg[1])
            continue

        _, li_inner = seg

        # Look at the next <p> segment to find margin-left → depth
        depth = None
        for j in range(idx + 1, min(len(segments), idx + 5)):
            if segments[j][0] == 'text':
                ml = re.search(r'margin-left:\s*([0-9.]+)(cm|px|mm)', segments[j][1])
                if ml:
                    val = float(ml.group(1))
                    if ml.group(2) == 'cm':
                        depth = int(val) - 1  # 1cm→0, 2cm→1, 3cm→2
                    break

        # Fallback: try to infer from OL nesting
        if depth is None or depth < 0:
            # Find the ol block text before this item
            prev_text_seg = None
            for j in range(idx - 1, -1, -1):
                if segments[j][0] == 'text':
                    prev_text_seg = segments[j][1]
                    break
            if prev_text_seg:
                ols_before = prev_text_seg.count('<ol ') + prev_text_seg.count('<ol>')
                depth = max(0, ols_before - 1)
            else:
                depth = 0
        depth = min(depth, 4)  # cap at 4

        # Track counter
        counters[depth] += 1
        for d in range(depth + 1, 6):
            counters[d] = 0

        margin = f'{depth + 1}cm'
        list_type = CSS_LIST_STYLE.get(depth, 'decimal')

        start_attr = ''
        if counters[depth] > 1:
            start_attr = f' start="{counters[depth]}"'

        result.append(
            f'<ol{start_attr} style="margin-left:{margin};padding-left:0;list-style-type:{list_type}">'
            f'<li>{li_inner}</li></ol>'
        )

    return ''.join(result)

# === HTTP Handler ===
class ConvertHandler(http.server.BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        # Quiet logging - only errors
        if '200' not in str(args):
            super().log_message(fmt, *args)

    def _send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def _send_cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.send_header('Access-Control-Max-Age', '86400')

    def do_OPTIONS(self):
        self.send_response(204)
        self._send_cors()
        self.end_headers()

    def do_GET(self):
        path = urlparse(self.path).path

        if path == '/health':
            soffice = find_soffice()
            self._send_json({
                'status': 'ok',
                'libreoffice': soffice or None,
                'libreoffice_available': soffice is not None,
            })
            return

        # Serve index.html and static files
        if path == '/' or path == '/index.html':
            self._serve_file('index.html', 'text/html')
            return

        # Serve other static files from the same directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        rel_path = path.lstrip('/')
        file_path = os.path.join(script_dir, rel_path)
        if os.path.isfile(file_path) and not rel_path.startswith('..'):
            ext = os.path.splitext(file_path)[1].lower()
            content_types = {
                '.html': 'text/html', '.css': 'text/css', '.js': 'application/javascript',
                '.json': 'application/json', '.png': 'image/png', '.jpg': 'image/jpeg',
                '.svg': 'image/svg+xml', '.ico': 'image/x-icon',
            }
            self._serve_file(file_path, content_types.get(ext, 'application/octet-stream'), True)
            return

        self._send_json({'error': 'not found'}, 404)

    def _serve_file(self, filepath, content_type, is_abs=False):
        if not is_abs:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            filepath = os.path.join(script_dir, filepath)
        try:
            with open(filepath, 'rb') as f:
                data = f.read()
            self.send_response(200)
            self.send_header('Content-Type', f'{content_type}; charset=utf-8')
            self.send_header('Content-Length', str(len(data)))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Cache-Control', 'no-cache')
            self.end_headers()
            self.wfile.write(data)
        except FileNotFoundError:
            self._send_json({'error': 'file not found'}, 404)

    def do_POST(self):
        path = urlparse(self.path).path

        if path != '/convert':
            self._send_json({'error': 'not found'}, 404)
            return

        content_length = int(self.headers.get('Content-Length', 0))
        if content_length > 50 * 1024 * 1024:  # 50MB limit
            self._send_json({'error': 'file too large (max 50MB)'}, 413)
            return

        content_type = self.headers.get('Content-Type', '')
        if 'multipart/form-data' not in content_type:
            self._send_json({'error': 'expected multipart/form-data'}, 400)
            return

        # Parse boundary
        boundary = None
        for part in content_type.split(';'):
            part = part.strip()
            if part.startswith('boundary='):
                boundary = part[9:].strip('"')
                break
        if not boundary:
            self._send_json({'error': 'missing boundary'}, 400)
            return

        # Read body
        body = self.rfile.read(content_length)

        # Parse multipart - find file data
        boundary_bytes = boundary.encode()
        parts = body.split(b'--' + boundary_bytes)

        file_data = None
        filename = None

        for part in parts:
            if not part or part.strip() == b'--' or part.strip() == b'':
                continue
            # Find headers
            header_end = part.find(b'\r\n\r\n')
            if header_end < 0:
                continue
            headers_raw = part[:header_end].decode('utf-8', errors='replace')
            file_body = part[header_end + 4:]
            # Remove trailing \r\n
            if file_body.endswith(b'\r\n'):
                file_body = file_body[:-2]

            if 'filename=' in headers_raw:
                # Extract filename
                fn_match = re.search(r'filename="([^"]+)"', headers_raw)
                if fn_match:
                    filename = fn_match.group(1)
                file_data = file_body
                break

        if not file_data or not filename:
            self._send_json({'error': 'no file found in upload'}, 400)
            return

        # Validate extension
        if not filename.lower().endswith(('.docx', '.doc')):
            self._send_json({'error': 'only .docx/.doc files are supported'}, 400)
            return

        # Check for old .doc format
        if file_data[:2] == b'\xd0\xcf':
            self._send_json({'error': 'old .doc format detected, please save as .docx'}, 400)
            return

        # Convert using LibreOffice
        with tempfile.TemporaryDirectory(prefix='docx2html_') as tmpdir:
            # Save uploaded file
            docx_path = os.path.join(tmpdir, filename)
            with open(docx_path, 'wb') as f:
                f.write(file_data)

            html_body, error = convert_with_libreoffice(docx_path, tmpdir)

            if error:
                self._send_json({'error': error, 'html': None}, 500)
                return

            self._send_json({
                'html': html_body,
                'filename': filename,
                'size': len(file_data),
                'html_size': len(html_body.encode('utf-8')),
            })

def main():
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8765

    soffice = find_soffice()
    print(f'=' * 50)
    print(f'  DOCX → HTML Converter Server')
    print(f'  Port: {port}')
    print(f'  LibreOffice: {soffice or "NOT FOUND"}')
    print(f'  URL: http://localhost:{port}')
    print(f'=' * 50)

    if not soffice:
        print()
        print('  WARNING: LibreOffice not detected!')
        print('  .docx conversion will not work until LibreOffice is installed.')
        print('  Download: https://www.libreoffice.org/download/')
        print()
        print('  You can still import .html files directly.')
        print()

    server = http.server.HTTPServer(('127.0.0.1', port), ConvertHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nServer stopped.')
        server.server_close()

if __name__ == '__main__':
    main()
