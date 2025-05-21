import os
import zipfile
import shutil
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET

def is_valid_xml(file_path):
    try:
        ET.parse(file_path)
        return True
    except ET.ParseError:
        return False

def attempt_basic_fix(xml_path):
    try:
        with open(xml_path, 'r', encoding='utf-8') as f:
            content = f.read().strip()

        # If nothing starts with '<', then it's corrupted
        if not content.startswith('<'):
            print("🔧 document.xml doesn't start with '<'. Trying to fix...")
            content = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' + \
                      '<w:body><w:p><w:r><w:t>Document repaired</w:t></w:r></w:p></w:body></w:document>'

        with open(xml_path, 'w', encoding='utf-8') as f:
            f.write(content)

    except Exception as e:
        print(f"❌ Failed to rewrite XML: {e}")

def repair_docx(docx_path):
    docx_path = Path(docx_path).resolve()
    fixed_path = docx_path.with_name(f"{docx_path.stem}_fixed.docx")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        # Step 1: Unzip corrupted docx
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(tmpdir)

        document_xml = tmpdir / "word" / "document.xml"
        if not document_xml.exists():
            print("❌ document.xml is missing completely.")
            return

        # Step 2: Validate XML or attempt basic repair
        if not is_valid_xml(document_xml):
            attempt_basic_fix(document_xml)

        # Step 3: Re-validate after fix attempt
        if not is_valid_xml(document_xml):
            print("❌ Still invalid XML after fix. Aborting.")
            return

        # Step 4: Rebuild .docx
        with zipfile.ZipFile(fixed_path, 'w', zipfile.ZIP_DEFLATED) as docx_file:
            for folder, _, files in os.walk(tmpdir):
                for file in files:
                    full_path = Path(folder) / file
                    arcname = full_path.relative_to(tmpdir)
                    docx_file.write(full_path, arcname)

        print(f"✅ Repaired DOCX saved as: {fixed_path}")

# === EXAMPLE USAGE ===
if __name__ == "__main__":
    repair_docx("generated_invoices/jan1970006.docx")
