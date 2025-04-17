import os
import unittest
from docx import Document
from src.save_formatting import extract_formatting_from_column

class TestDocxMerge(unittest.TestCase):
    INPUT_PATH = "data/input_files/250403_三井海洋開発様_P.13_ビジネスモデル-ja-en-D.docx"
    OUTPUT_FOLDER = "data/output_files"
    OUTPUT_PATH = os.path.join(
        OUTPUT_FOLDER,
        f"{os.path.splitext(os.path.basename(INPUT_PATH))[0]}_merged.docx"
    )

    # 3rd table in input, 1st table in output
    INPUT_TABLE_INDEX = 3
    OUTPUT_TABLE_INDEX = 0

    # map input‐table cols → output‐table cols
    COL_MAP = {
        3: 1,  # input col 3 → output col 1 (Source/Japanese)
        5: 2,  # input col 5 → output col 2 (Target/English)
        6: 3,  # input col 6 → output col 3 (Match)
        7: 4   # input col 7 → output col 4 (Comment)
    }

    def setUp(self):
        self.input_doc = Document(self.INPUT_PATH)
        self.output_doc = Document(self.OUTPUT_PATH)

        self.in_table  = self.input_doc.tables[self.INPUT_TABLE_INDEX]
        self.out_table = self.output_doc.tables[self.OUTPUT_TABLE_INDEX]

        self.col_map = TestDocxMerge.COL_MAP

    def test_row_count(self):
        self.assertEqual(
            len(self.in_table.rows) + 1, # Input table does not have the header row
            len(self.out_table.rows),
            f"Row count mismatch: input has {len(self.in_table.rows)}, "
            f"output has {len(self.out_table.rows)}"
        )

    def test_cell_texts_mirror(self):
        for row_idx, (in_row, out_row) in enumerate(
            zip(self.in_table.rows, self.out_table.rows)
        ):
            for in_col, out_col in self.col_map.items():
                in_text  = in_row.cells[in_col].text.strip()
                out_text = out_row.cells[out_col].text.strip()
                self.assertEqual(
                    in_text, out_text,
                    f"Text mismatch at row {row_idx}, input col {in_col}, output col {out_col}"
                )

    def test_formatting_info_mirror(self):
        # pull run‐level formatting for just those cols
        in_fmt = extract_formatting_from_column(
            self.input_doc,  self.INPUT_TABLE_INDEX, list(self.col_map.keys())
        )
        out_fmt = extract_formatting_from_column(
            self.output_doc, self.OUTPUT_TABLE_INDEX, list(self.col_map.values())
        )

        for row_idx in in_fmt:
            for in_col, out_col in self.col_map.items():
                runs_in  = in_fmt[row_idx][in_col]
                runs_out = out_fmt[row_idx][out_col]
                self.assertEqual(
                    runs_in, runs_out,
                    f"Formatting mismatch at row {row_idx}, "
                    f"input col {in_col}, output col {out_col}"
                )

if __name__ == "__main__":
    unittest.main()
