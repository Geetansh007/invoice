import os
import shutil
import zipfile
import copy
from io import BytesIO
from typing import List, Dict, Any
import itertools
import xml.etree.ElementTree as ET


class InvoiceDocGeneratorXML:
    """
    Generate invoice DOCX files by patching the underlying XML of a Word template
    ("Invoice format.docx"). Styling is preserved perfectly—even when inserting an
    arbitrary number of line‑items—by cloning the first data row.

    **2025‑05‑20 update**
    • Placeholders can now be split across several <w:t> runs (e.g. "Col I"),
      so replacement works no matter how Word internally chunks the text.
    • The same paragraph‑level strategy is applied to the «Self generate …» logic.
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
        "Col I": "Address",  # now handled even when the run is split
    }

    _SELF_GEN_ORDER = (
        "Invoice Number",
        None,  # keeps the literal text "Self generate" the second time
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

    # ────────────────────────────────────────────────────────────────────
    # public
    # ────────────────────────────────────────────────────────────────────
    def __init__(self, template_path: str, output_dir: str, data_list: List[Dict[str, Any]]):
        self.template_path = template_path
        self.output_dir = output_dir
        self.data_list = data_list
        os.makedirs(self.output_dir, exist_ok=True)

    def generate_documents(self) -> None:
        for payload in self.data_list:
            file_name = payload.get("Invoice Number")
            target_docx = os.path.join(self.output_dir, f"{file_name}.docx")
            shutil.copy2(self.template_path, target_docx)
            self._patch_docx(target_docx, payload)
            print(f"✔  Saved → {target_docx}")

    # ────────────────────────────────────────────────────────────────────
    # internal helpers
    # ────────────────────────────────────────────────────────────────────
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
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    # ────────────────────────────────────────────────────────────────────
    # text replacement  (paragraph‑level, handles split runs)
    # ────────────────────────────────────────────────────────────────────
    def _replace_placeholders(self, root, data):
        """Replace every placeholder in _PLACEHOLDERS, even if split across runs."""
        for p in root.iterfind(".//w:p", self._NS):
            runs = list(p.iterfind(".//w:t", self._NS))
            if not runs:
                continue
            para_text = "".join(t.text or "" for t in runs)
            for ph, field in self._PLACEHOLDERS.items():
                para_text = para_text.replace(ph, str(data.get(field, "")))
            # write back into the first run and blank the rest (keeps formatting)
            runs[0].text = para_text
            for t in runs[1:]:
                t.text = ""

    def _replace_self_generate(self, root, data):
        """Handle the special "Self generate" tokens with paragraph‑level safety."""
        counter = 0
        for p in root.iterfind(".//w:p", self._NS):
            runs = list(p.iterfind(".//w:t", self._NS))
            if not runs:
                continue
            para_text = "".join(t.text or "" for t in runs)
            if "Self generate" not in para_text:
                continue
            parts = para_text.split("Self generate")
            rebuilt = [parts[0]]
            for _ in range(1, len(parts)):
                counter += 1
                field = self._SELF_GEN_ORDER[counter - 1] if counter - 1 < len(self._SELF_GEN_ORDER) else None
                rebuilt.append(str(data.get(field, "")) if field else "Self generate")
                rebuilt.append(parts[_])
            new_text = "".join(rebuilt)
            runs[0].text = new_text
            for t in runs[1:]:
                t.text = ""

    # ────────────────────────────────────────────────────────────────────
    # table population
    # ────────────────────────────────────────────────────────────────────
    def _populate_invoice_table(self, root, data):
        tbl = self._find_invoice_table(root)
        if tbl is None:
            return

        descriptions = data.get("Description", [])
        amounts = data.get("Amount", [])

        descriptions = descriptions if isinstance(descriptions, list) else [descriptions]
        amounts = amounts if isinstance(amounts, list) else [amounts]

        if len(descriptions) != len(amounts):
            print("⚠  Description / Amount length mismatch – skipped table fill.")
            return

        rows = tbl.findall("w:tr", self._NS)
        if len(rows) < 2:
            return

        template_row = rows[1]
        for r in rows[1:]:
            tbl.remove(r)

        # ── 1) find every ID already used anywhere in document.xml
        used_ids = self._collect_existing_ids(root)
        start = max(used_ids) + 1 if used_ids else 1
        self._id_counter = itertools.count(start)

        # ── 2) add rows
        for idx, (desc, amt) in enumerate(zip(descriptions, amounts), start=1):
            new_row = copy.deepcopy(template_row)
            self._fix_unique_ids(new_row)
            tcs = new_row.findall("w:tc", self._NS)
            self._set_first_run_text(tcs[0], str(idx))
            self._set_first_run_text(tcs[1], str(desc))
            self._set_first_run_text(tcs[4], str(amt))
            tbl.append(new_row)

        # restore footer rows if the template had any
        if len(rows) > len(descriptions) + 1:
            for r in rows[len(descriptions) + 1:]:
                tbl.append(r)

    # ────────────────────────────────────────────────────────────────────
    # unique‑ID handling
    # ────────────────────────────────────────────────────────────────────
    def _collect_existing_ids(self, root):
        ids = set()
        # bookmarkStart / bookmarkEnd
        for bm in root.iterfind(".//w:bookmarkStart", self._NS):
            ids.add(int(bm.get(f"{{{self._NS['w']}}}id", "0")))
        for bm in root.iterfind(".//w:bookmarkEnd", self._NS):
            ids.add(int(bm.get(f"{{{self._NS['w']}}}id", "0")))
        # drawings (docPr id)
        for docPr in root.iterfind(
            ".//wp:docPr",
            {'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'}
        ):
            ids.add(int(docPr.get("id", "0")))
        # anchorId is hex
        for anchor in root.iterfind(
            ".//wp14:anchor",
            {'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'}
        ):
            ids.add(int(anchor.get(
                "{http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing}anchorId",
                "0"
            ), 16))
        return ids

    def _get_next_id(self) -> int:
        return next(self._id_counter)

    def _fix_unique_ids(self, row):
        for bm in row.iterfind(".//w:bookmarkStart", self._NS):
            bm.set(f"{{{self._NS['w']}}}id", str(self._get_next_id()))
        for bm in row.iterfind(".//w:bookmarkEnd", self._NS):
            bm.set(f"{{{self._NS['w']}}}id", str(self._get_next_id()))
        for docPr in row.iterfind(
            ".//wp:docPr",
            {'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'}
        ):
            docPr.set("id", str(self._get_next_id()))
        for anchor in row.iterfind(
            ".//wp14:anchor",
            {'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'}
        ):
            anchor.set(
                "{http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing}anchorId",
                f"{self._get_next_id():08X}"
            )

    # ────────────────────────────────────────────────────────────────────
    # misc helpers
    # ────────────────────────────────────────────────────────────────────
    def _find_invoice_table(self, root):
        for tbl in root.iterfind(".//w:tbl", self._NS):
            first_row = tbl.find("w:tr", self._NS)
            if first_row is None:
                continue
            hdrs = [
                "".join(t.text or "" for t in c.iterfind(".//w:t", self._NS)).strip().lower()
                for c in first_row.findall("w:tc", self._NS)
            ]
            if all(exp in hdrs for exp in self._INVOICE_TABLE_HEADERS):
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
