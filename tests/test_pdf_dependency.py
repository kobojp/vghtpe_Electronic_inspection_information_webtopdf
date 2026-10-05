import unittest

import fitz


class PdfDependencyTests(unittest.TestCase):
    def test_search_and_extract_pdf_page(self):
        with fitz.open() as source:
            page = source.new_page()
            page.insert_text((72, 72), "PDF dependency smoke test")
            pdf_bytes = source.tobytes()

        with fitz.open(stream=pdf_bytes, filetype="pdf") as source:
            self.assertTrue(source[0].search_for("dependency"))
            with fitz.open() as extracted:
                extracted.insert_pdf(source, from_page=0, to_page=0)
                output = extracted.tobytes()

        with fitz.open(stream=output, filetype="pdf") as result:
            self.assertEqual(result.page_count, 1)
            self.assertIn("PDF dependency smoke test", result[0].get_text())


if __name__ == "__main__":
    unittest.main()
