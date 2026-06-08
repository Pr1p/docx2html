"""
LibreOffice DOCX→HTML 批量转换脚本
使用方法:
    python lo_convert.py <input.docx或目录> [output_dir]

依赖: 需要安装 LibreOffice (soffice 在 PATH 中或自动检测常见路径)
"""
import subprocess
import sys
import os
import shutil
import glob
import re

# LibreOffice 常见安装路径
SOFFICE_PATHS = [
    r'C:\Program Files\LibreOffice\program\soffice.exe',
    r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
    '/usr/bin/soffice',
    '/usr/bin/libreoffice',
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
]

def find_soffice():
    # Check PATH first
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if soffice:
        return soffice
    # Check common paths
    for p in SOFFICE_PATHS:
        if os.path.isfile(p):
            return p
    return None

def convert_docx_to_html(docx_path, output_dir=None):
    """Convert a single .docx to .html using LibreOffice"""
    soffice = find_soffice()
    if not soffice:
        print("ERROR: LibreOffice not found. Please install from https://www.libreoffice.org/")
        sys.exit(1)

    docx_path = os.path.abspath(docx_path)
    if not os.path.isfile(docx_path):
        print(f"ERROR: File not found: {docx_path}")
        sys.exit(1)

    if output_dir is None:
        output_dir = os.path.dirname(docx_path)
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    print(f"Converting: {os.path.basename(docx_path)}")
    print(f"  soffice: {soffice}")
    print(f"  output: {output_dir}")

    cmd = [
        soffice,
        '--headless',
        '--convert-to', 'html',
        '--outdir', output_dir,
        docx_path,
    ]

    result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)

    if result.returncode != 0:
        print(f"ERROR: Conversion failed (exit code {result.returncode})")
        if result.stderr:
            print(result.stderr)
        return None

    # LibreOffice outputs with same basename but .html extension
    basename = os.path.splitext(os.path.basename(docx_path))[0] + '.html'
    output_path = os.path.join(output_dir, basename)

    if os.path.isfile(output_path):
        # Clean the HTML - extract body content and remove Word artifacts
        with open(output_path, 'r', encoding='utf-8', errors='replace') as f:
            html = f.read()

        # Extract body
        body_match = re.search(r'<body[^>]*>([\s\S]*?)</body>', html, re.IGNORECASE)
        body = body_match.group(1).strip() if body_match else html

        # Clean up LibreOffice HTML artifacts
        body = re.sub(r'<\?xml[^?]*\?>', '', body)
        body = re.sub(r'<\/?\w+:\w+[^>]*>', '', body)  # namespace tags
        body = re.sub(r'<!--\[if[\s\S]*?<!\[endif\]-->', '', body)  # conditional comments
        body = re.sub(r'\s+class="(?:Western|TE|SD|Caption|Title|Heading|Text|List|Frame|Table)[^"]*"', '', body)
        body = re.sub(r'\s+lang="[^"]*"', '', body)
        body = re.sub(r'<span\s*>', '', body)
        body = re.sub(r'</span>', '', body)

        # Write cleaned version
        clean_path = os.path.join(output_dir, basename)
        with open(clean_path, 'w', encoding='utf-8') as f:
            f.write(body)

        print(f"  Done: {clean_path} ({len(body)} chars)")
        return clean_path
    else:
        print(f"WARNING: Output file not found at expected path: {output_path}")
        # Try to find it
        html_files = glob.glob(os.path.join(output_dir, '*.html'))
        if html_files:
            print(f"  Found: {html_files[0]}")
            return html_files[0]
        return None

def main():
    if len(sys.argv) < 2:
        print("Usage: python lo_convert.py <input.docx> [output_dir]")
        print("       python lo_convert.py <input_dir> [output_dir]")
        sys.exit(1)

    input_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None

    if os.path.isdir(input_path):
        # Convert all .docx in directory
        docx_files = glob.glob(os.path.join(input_path, '*.docx'))
        if not docx_files:
            print(f"No .docx files found in: {input_path}")
            sys.exit(1)
        for f in docx_files:
            convert_docx_to_html(f, output_dir)
    else:
        convert_docx_to_html(input_path, output_dir)

if __name__ == '__main__':
    main()
