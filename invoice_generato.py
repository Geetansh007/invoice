# invoice_xml_generator.py
import os
import re
import shutil
import zipfile
from io import BytesIO
from typing import List, Dict, Any
from datetime import datetime
import xml.etree.ElementTree as ET


class InvoiceDocGeneratorXML:
    """
    Same public interface as your original InvoiceDocGenerator,
    but all work is done by patching the XML inside the DOCX.
    """

    # ------------------------------------------------------------------ #
    # ❶  Configuration                                                   #
    # ------------------------------------------------------------------ #
    _NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    ET.register_namespace("w", _NS["w"])          # keep the familiar <w:..> tags

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
        "Invoice Number",          # 1st “Self generate”
        None,                      # 2nd  – leave literal text
        "Total Amount",            # 3rd
        "Total Amount in Words",   # 4th
    )

    _INVOICE_TABLE_HEADERS = [
        "sr. no.",
        "description",
        "unit",
        "unit/ price",
        "amount",
    ]

    # ------------------------------------------------------------------ #
    # ❷  Construction                                                    #
    # ------------------------------------------------------------------ #
    def __init__(
        self,
        template_path: str,
        output_dir: str,
        data_list: List[Dict[str, Any]],
    ):
        self.template_path = template_path
        self.output_dir = output_dir
        self.data_list = data_list
        os.makedirs(self.output_dir, exist_ok=True)

    # ------------------------------------------------------------------ #
    # ❸  Public entry-point                                              #
    # ------------------------------------------------------------------ #
    def generate_documents(self) -> None:
        for payload in self.data_list:
            file_name = (
                payload.get("Invoice Number")
                or payload.get("Beneficiary Name")
                or f"invoice_{datetime.now():%Y%m%d%H%M%S}"
            )
            target_docx = os.path.join(self.output_dir, f"{file_name}.docx")

            # 1. copy the template
            shutil.copy2(self.template_path, target_docx)

            # 2. patch document.xml in-place
            self._patch_docx(target_docx, payload)

            print(f"✔  Saved → {target_docx}")

    # ------------------------------------------------------------------ #
    # ❹  ZIP-level work                                                  #
    # ------------------------------------------------------------------ #
    def _patch_docx(self, docx_path: str, data: Dict[str, Any]) -> None:
        with zipfile.ZipFile(docx_path, mode="r") as zin:
            buf = BytesIO()
            with zipfile.ZipFile(
                buf, mode="w", compression=zipfile.ZIP_DEFLATED
            ) as zout:
                for item in zin.infolist():
                    if item.filename == "word/document.xml":
                        xml_in = zin.read(item.filename)
                        xml_out = self._transform_document_xml(xml_in, data)
                        zout.writestr(item, xml_out)
                    else:
                        zout.writestr(item, zin.read(item.filename))
        # overwrite
        with open(docx_path, "wb") as f:
            f.write(buf.getvalue())

    # ------------------------------------------------------------------ #
    # ❺  XML-level transformations                                       #
    # ------------------------------------------------------------------ #
    def _transform_document_xml(
        self, xml_bytes: bytes, data: Dict[str, Any]
    ) -> bytes:
        root = ET.fromstring(xml_bytes)

        # (a) ordinary Col-placeholders inside any run
        self._replace_placeholders(root, data)

        # (b) sequential “Self generate”
        self._replace_self_generate(root, data)

        # (c) invoice-items table rows
        self._populate_invoice_table(root, data)

        return ET.tostring(root, encoding="utf-8")

    # ---- a) placeholder replacement ---------------------------------- #
    def _replace_placeholders(self, root, data):
        for t in root.iterfind(".//w:t", self._NS):
            txt = t.text or ""
            for ph, field in self._PLACEHOLDERS.items():
                if ph in txt:
                    val = str(data.get(field, ""))
                    t.text = txt.replace(ph, val)

    # ---- b) “Self generate” sequence ---------------------------------- #
    def _replace_self_generate(self, root, data):
        counter = 0
        for t in root.iterfind(".//w:t", self._NS):
            txt = t.text or ""
            if "Self generate" not in txt:
                continue

            parts = txt.split("Self generate")
            rebuilt = [parts[0]]

            for _ in range(1, len(parts)):
                # which occurrence is this, *in the whole document*?
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

    # ---- c) fill the invoice items table ------------------------------ #
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
        if len(rows) <= 1:
            return                                           # template too small

        for idx, (desc, amt) in enumerate(zip(descriptions, amounts), start=1):
            if idx >= len(rows):
                print("⚠  Not enough blank rows in template – remaining items ignored.")
                break

            tc_srno, tc_desc, *_, tc_amount = rows[idx].findall("w:tc", self._NS)
            self._set_cell_text(tc_srno, str(idx))
            self._set_cell_text(tc_desc, str(desc))
            self._set_cell_text(tc_amount, str(amt))

    def _find_invoice_table(self, root):
        for tbl in root.iterfind(".//w:tbl", self._NS):
            first_row = tbl.find("w:tr", self._NS)
            if first_row is None:
                continue
            cells = first_row.findall("w:tc", self._NS)
            hdrs = [
                "".join(t.text or "" for t in c.iterfind(".//w:t", self._NS))
                .strip()
                .lower()
                for c in cells
            ]
            if all(any(h == exp for h in hdrs) for exp in self._INVOICE_TABLE_HEADERS):
                return tbl
        return None

    # ------------------------------------------------------------------ #
    # ❻  Tiny helpers                                                    #
    # ------------------------------------------------------------------ #
    def _set_cell_text(self, tc, new_text: str):
        """
        Overwrite **only** the first <w:t> inside the cell; leave styling alone.
        """
        first_t = tc.find(".//w:t", self._NS)
        if first_t is not None:
            first_t.text = new_text
            # clear any additional <w:t> siblings so no leftover text shows
            for extra_t in tc.findall(".//w:t", self._NS)[1:]:
                extra_t.text = ""
