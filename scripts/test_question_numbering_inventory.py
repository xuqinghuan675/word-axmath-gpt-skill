"""Exercise direct and inherited list numbering without modifying a document."""
from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

from inspect_question_numbering import inspect, sha256

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class NumberingInventoryTests(unittest.TestCase):
    def test_detects_list_style_fields_and_outline_with_override(self):
        document = f'''<w:document xmlns:w="{W}"><w:body>
          <w:p><w:r><w:t>4. 普通文本题号</w:t></w:r></w:p>
          <w:p><w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>直接列表</w:t></w:r></w:p>
          <w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr><w:r><w:t>继承列表</w:t></w:r></w:p>
          <w:p><w:pPr><w:pStyle w:val="Child"/><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>禁用列表</w:t></w:r></w:p>
          <w:p><w:r><w:instrText> LI</w:instrText></w:r><w:r><w:instrText>STNUM </w:instrText></w:r></w:p>
          <w:p><w:fldSimple w:instr=" SEQ Question "><w:r><w:t>题号域</w:t></w:r></w:fldSimple></w:p>
          <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>标题</w:t></w:r></w:p>
        </w:body></w:document>'''
        styles = f'''<w:styles xmlns:w="{W}">
          <w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>
          <w:style w:type="paragraph" w:styleId="List"><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr></w:style>
          <w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="List"/></w:style>
          <w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
        </w:styles>'''
        header = f'<w:hdr xmlns:w="{W}"><w:p><w:fldSimple w:instr=" AUTONUM "/></w:p></w:hdr>'
        with tempfile.TemporaryDirectory() as directory:
            docx = Path(directory) / "questions.docx"
            with zipfile.ZipFile(docx, "w") as package:
                package.writestr("word/document.xml", document)
                package.writestr("word/styles.xml", styles)
                package.writestr("word/header1.xml", header)
            before = sha256(docx)
            report = inspect(docx)
            self.assertEqual(before, sha256(docx))
            body = {x["paragraph"]: x for x in report["candidates"] if x["story"] == "word/document.xml"}
            self.assertEqual(set(body), {2, 3, 5, 6, 7})
            self.assertEqual(body[2]["numbering_origin"], "paragraph")
            self.assertEqual(body[3]["numbering_origin"], "style:List")
            self.assertEqual(body[5]["numbering_fields"], ["LISTNUM"])
            self.assertEqual(body[6]["numbering_fields"], ["SEQ"])
            self.assertEqual(body[7]["outline_level"], "0")
            self.assertTrue(any(x["story"] == "word/header1.xml" and x["numbering_fields"] == ["AUTONUM"] for x in report["candidates"]))


if __name__ == "__main__":
    unittest.main()
