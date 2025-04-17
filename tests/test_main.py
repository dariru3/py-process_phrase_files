import unittest
from docx import Document
from src.save_formatting import extract_formatting_from_column

class TestDocxMerge(unittest.TestCase):
    INPUT_PATH = "data/input_files/250403_三井海洋開発様_P.13_ビジネスモデル-ja-en-D.docx"
    OUTPUT_PATH = "data/output_files/250403_三井海洋開発様_P.13_ビジネスモデル-ja-en-D_merged.docx"
    INPUT_TABLE_INDEX = 2   # 3rd table in the input
    OUTPUT_TABLE_INDEX = 0  # 1st table in the output

    def setUp(self):
        self.input_doc = Document(self.INPUT_PATH)
        self.output_doc = Document(self.OUTPUT_PATH)
        # point at the tables we actually care about
        self.input_table = self.input_doc.tables[self.INPUT_TABLE_INDEX]
        self.output_table = self.output_doc.tables[self.OUTPUT_TABLE_INDEX]

        # assume these two tables have the same number of columns
        self.n_cols = len(self.input_table.columns)
        self.col_nums = list(range(self.n_cols))

    def test_row_count(self):
        in_rows  = len(self.input_table.rows)
        out_rows = len(self.output_table.rows)
        self.assertEqual(
            in_rows, out_rows,
            f"Row count mismatch: input has {in_rows}, output has {out_rows}"
        )

    def test_cell_texts_mirror(self):
        for idx, (in_row, out_row) in enumerate(zip(self.input_table.rows,
                                                   self.output_table.rows), start=1):
            in_texts  = [cell.text.strip() for cell in in_row.cells]
            out_texts = [cell.text.strip() for cell in out_row.cells]
            self.assertListEqual(
                in_texts, out_texts,
                f"Row {idx} text mismatch:\n  input:  {in_texts}\n  output: {out_texts}"
            )

    def test_formatting_info_mirror(self):
        # extract run‐level formatting from both tables
        in_fmt  = extract_formatting_from_column(
            self.input_doc, self.INPUT_TABLE_INDEX, self.col_nums
        )
        out_fmt = extract_formatting_from_column(
            self.output_doc, self.OUTPUT_TABLE_INDEX, self.col_nums
        )
        # compare the nested dicts straight‑up
        self.assertDictEqual(
            in_fmt, out_fmt,
            "Run‑level formatting info differs between input (3rd table) and output (1st table)"
        )

if __name__ == "__main__":
    unittest.main()
