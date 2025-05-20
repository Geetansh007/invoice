import os
import shutil
import zipfile
import copy
from io import BytesIO
from typing import List, Dict, Any
from datetime import datetime
import xml.etree.ElementTree as ET


class InvoiceDocGeneratorXML:
    """
    Generate invoice DOCX files by patching the underlying XML of a Word template
    ("Invoice format.docx").  Styling is preserved perfectly—even when inserting an
    arbitrary number of line‑items—by cloning the first data row.
    """

    _NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    ET.register_namespace("w", _NS["w"])

    _PLACEHOLDERS = {
        "Col A": "Date of Entry",
        "Col B": "Amount",
        "Col C": "Description",
        "Col D": "Beneficiary Name",
        "Col E": "Bank Name",
        "Col F": "Bank Account No",
        "Col G": "IFSC/SWIFT Code",
        "Col H": "IBAN No",
        "Col I": "Address",
    }

    _SELF_GEN_ORDER = (
        "Invoice Number",
        None,
        "Total Amount",
        "Total Amount in Words",
    )

    _INVOICE_TABLE_HEADERS = [
        "sr. no.",
        "description",
        "unit",
        "unit/ price",
        "amount",
    ]

    def __init__(self, template_path: str, output_dir: str, data_list: List[Dict[str, Any]]):
        self.template_path = template_path
        self.output_dir = output_dir
        self.data_list = data_list
        os.makedirs(self.output_dir, exist_ok=True)

    def generate_documents(self) -> None:
        for payload in self.data_list:
            file_name = (
                payload.get("Invoice Number")
                or payload.get("Beneficiary Name")
                or f"invoice_{datetime.now():%Y%m%d%H%M%S}"
            )
            target_docx = os.path.join(self.output_dir, f"{file_name}.docx")
            shutil.copy2(self.template_path, target_docx)
            self._patch_docx(target_docx, payload)
            print(f"✔  Saved → {target_docx}")

    def _patch_docx(self, docx_path: str, data: Dict[str, Any]) -> None:
        with zipfile.ZipFile(docx_path, mode="r") as zin:
            buf = BytesIO()
            with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == "word/document.xml":
                        xml_in = zin.read(item.filename)
                        xml_out = self._transform_document_xml(xml_in, data)
                        zout.writestr(item, xml_out)
                    else:
                        zout.writestr(item, zin.read(item.filename))
        with open(docx_path, "wb") as f:
            f.write(buf.getvalue())

    def _transform_document_xml(self, xml_bytes: bytes, data: Dict[str, Any]) -> bytes:
        root = ET.fromstring(xml_bytes)
        self._replace_placeholders(root, data)
        self._replace_self_generate(root, data)
        self._populate_invoice_table(root, data)
        return ET.tostring(root, encoding="utf-8")

    def _replace_placeholders(self, root, data):
        for t in root.iterfind(".//w:t", self._NS):
            txt = t.text or ""
            for ph, field in self._PLACEHOLDERS.items():
                if ph in txt:
                    val = str(data.get(field, ""))
                    t.text = txt.replace(ph, val)

    def _replace_self_generate(self, root, data):
        counter = 0
        for t in root.iterfind(".//w:t", self._NS):
            txt = t.text or ""
            if "Self generate" not in txt:
                continue
            parts = txt.split("Self generate")
            rebuilt = [parts[0]]
            for _ in range(1, len(parts)):
                counter += 1
                repl_field = (
                    self._SELF_GEN_ORDER[counter - 1]
                    if counter - 1 < len(self._SELF_GEN_ORDER)
                    else None
                )
                replacement = (
                    str(data.get(repl_field, "")) if repl_field else "Self generate"
                )
                rebuilt.append(replacement)
                rebuilt.append(parts[_])
            t.text = "".join(rebuilt)

    def _populate_invoice_table(self, root, data):
        tbl = self._find_invoice_table(root)
        if tbl is None:
            return

        descriptions = data.get("Description", [])
        amounts = data.get("Amount", [])

        if not isinstance(descriptions, list):
            descriptions = [descriptions]
        if not isinstance(amounts, list):
            amounts = [amounts]

        if len(descriptions) != len(amounts):
            print("⚠  Description / Amount length mismatch – skipped table fill.")
            return

        rows = tbl.findall("w:tr", self._NS)
        if len(rows) < 2:
            return

        template_row = rows[1]
        for r in rows[1:]:
            tbl.remove(r)

        for idx, (desc, amt) in enumerate(zip(descriptions, amounts), start=1):
            new_row = copy.deepcopy(template_row)
            tcs = new_row.findall("w:tc", self._NS)
            self._set_first_run_text(tcs[0], str(idx))
            self._set_first_run_text(tcs[1], str(desc))
            self._set_first_run_text(tcs[4], str(amt))
            tbl.append(new_row)

        # Restore last total row from original table if it exists
        if len(rows) > len(descriptions) + 1:
            for r in rows[len(descriptions) + 1:]:
                tbl.append(r)

    def _find_invoice_table(self, root):
        for tbl in root.iterfind(".//w:tbl", self._NS):
            first_row = tbl.find("w:tr", self._NS)
            if first_row is None:
                continue
            cells = first_row.findall("w:tc", self._NS)
            hdrs = [
                "".join(t.text or "" for t in c.iterfind(".//w:t", self._NS)).strip().lower()
                for c in cells
            ]
            if all(any(h == exp for h in hdrs) for exp in self._INVOICE_TABLE_HEADERS):
                return tbl
        return None

    def _set_first_run_text(self, tc, new_text: str):
        first_r = tc.find(".//w:r", self._NS)
        if first_r is None:
            return
        t = first_r.find("w:t", self._NS)
        if t is None:
            t = ET.SubElement(first_r, f"{{{self._NS['w']}}}t")
        t.text = new_text
