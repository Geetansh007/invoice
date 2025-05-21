import os
import shutil
import zipfile
import copy
from io import BytesIO
from typing import List, Dict, Any
import itertools
from lxml import etree as ET  # Use lxml instead of xml.etree.ElementTree


class InvoiceDocGeneratorXML:
    _NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

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
            file_name = payload.get("Invoice Number")
            target_docx = os.path.join(self.output_dir, f"{file_name}.docx")
            shutil.copy2(self.template_path, target_docx)
            self._patch_docx(target_docx, payload)
            print(f"\u2714  Saved → {target_docx}")

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
        parser = ET.XMLParser(recover=True)
        root = ET.fromstring(xml_bytes, parser=parser)
        self._replace_placeholders(root, data)
        self._replace_self_generate(root, data)
        self._populate_invoice_table(root, data)
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    def _replace_placeholders(self, root, data):
        for p in root.xpath(".//w:p", namespaces=self._NS):
            runs = p.xpath(".//w:t", namespaces=self._NS)
            if not runs:
                continue
            para_text = "".join(t.text or "" for t in runs)
            for ph, field in self._PLACEHOLDERS.items():
                para_text = para_text.replace(ph, str(data.get(field, "")))
            runs[0].text = para_text
            for t in runs[1:]:
                t.text = ""

    def _replace_self_generate(self, root, data):
        counter = 0
        for p in root.xpath(".//w:p", namespaces=self._NS):
            runs = p.xpath(".//w:t", namespaces=self._NS)
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

    def _populate_invoice_table(self, root, data):
        tbl = self._find_invoice_table(root)
        if tbl is None:
            return

        descriptions = data.get("Description", [])
        amounts = data.get("Amount", [])

        descriptions = descriptions if isinstance(descriptions, list) else [descriptions]
        amounts = amounts if isinstance(amounts, list) else [amounts]

        if len(descriptions) != len(amounts):
            print("\u26a0  Description / Amount length mismatch – skipped table fill.")
            return

        rows = tbl.xpath("w:tr", namespaces=self._NS)
        if len(rows) < 2:
            return

        template_row = rows[1]
        for r in rows[1:]:
            tbl.remove(r)

        used_ids = self._collect_existing_ids(root)
        start = max(used_ids) + 1 if used_ids else 1
        self._id_counter = itertools.count(start)

        for idx, (desc, amt) in enumerate(zip(descriptions, amounts), start=1):
            new_row = copy.deepcopy(template_row)
            self._fix_unique_ids(new_row)
            tcs = new_row.xpath("w:tc", namespaces=self._NS)
            self._set_first_run_text(tcs[0], str(idx))
            self._set_first_run_text(tcs[1], str(desc))
            self._set_first_run_text(tcs[4], str(amt))
            tbl.append(new_row)

        if len(rows) > len(descriptions) + 1:
            for r in rows[len(descriptions) + 1:]:
                tbl.append(r)

    def _collect_existing_ids(self, root):
        ids = set()
        for bm in root.xpath(".//w:bookmarkStart", namespaces=self._NS):
            ids.add(int(bm.get("w:id", "0")))
        for bm in root.xpath(".//w:bookmarkEnd", namespaces=self._NS):
            ids.add(int(bm.get("w:id", "0")))
        for docPr in root.xpath(".//wp:docPr", namespaces={"wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}):
            ids.add(int(docPr.get("id", "0")))
        for anchor in root.xpath(".//wp14:anchor", namespaces={"wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"}):
            ids.add(int(anchor.get("wp14:anchorId", "0"), 16))
        return ids

    def _get_next_id(self) -> int:
        return next(self._id_counter)

    def _fix_unique_ids(self, row):
        for bm in row.xpath(".//w:bookmarkStart", namespaces=self._NS):
            bm.set("w:id", str(self._get_next_id()))
        for bm in row.xpath(".//w:bookmarkEnd", namespaces=self._NS):
            bm.set("w:id", str(self._get_next_id()))
        for docPr in row.xpath(".//wp:docPr", namespaces={"wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}):
            docPr.set("id", str(self._get_next_id()))
        for anchor in row.xpath(".//wp14:anchor", namespaces={"wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"}):
            anchor.set("wp14:anchorId", f"{self._get_next_id():08X}")

    def _find_invoice_table(self, root):
        for tbl in root.xpath(".//w:tbl", namespaces=self._NS):
            first_row = tbl.find("w:tr", namespaces=self._NS)
            if first_row is None:
                continue
            hdrs = [
                "".join(t.text or "" for t in c.xpath(".//w:t", namespaces=self._NS)).strip().lower()
                for c in first_row.xpath("w:tc", namespaces=self._NS)
            ]
            if all(exp in hdrs for exp in self._INVOICE_TABLE_HEADERS):
                return tbl
        return None

    def _set_first_run_text(self, tc, new_text: str):
        first_r = tc.find(".//w:r", namespaces=self._NS)
        if first_r is None:
            return
        t = first_r.find("w:t", namespaces=self._NS)
        if t is None:
            t = ET.SubElement(first_r, f"{{{self._NS['w']}}}t")
        t.text = new_text
