import os
import tempfile
import unittest

from src.process_mxliff import parse_mxliff_to_df


class TestMxliffParsing(unittest.TestCase):
    def test_empty_source_text(self):
        mxliff_content = """<?xml version="1.0" encoding="UTF-8"?>
<xliff xmlns="urn:oasis:names:tc:xliff:document:1.2" version="1.2">
  <file>
    <body>
      <trans-unit id="1">
        <source/>
        <target>target text</target>
      </trans-unit>
    </body>
  </file>
</xliff>
"""
        with tempfile.NamedTemporaryFile("w+", suffix=".mxliff", delete=False) as tmp:
            tmp.write(mxliff_content)
            tmp_path = tmp.name

        try:
            df = parse_mxliff_to_df(tmp_path)
        finally:
            os.remove(tmp_path)

        self.assertEqual(df.loc[0, "ID"], "1")
        self.assertEqual(df.loc[0, "Source"], "")
        self.assertEqual(df.loc[0, "Target"], "target text")


if __name__ == "__main__":
    unittest.main()
