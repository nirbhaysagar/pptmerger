"""
test_suite.py - Automated Verification Suite for PPT Unify
Tests:
1. File validation logic (valid, corrupted, zero-byte)
2. Sequential merge with text, shapes, and pictures
3. Interleaved merge ("by line")
4. PDF export engine
5. High-load stress test with 100 PPTX files error-free
"""

import os
import io
import sys
import time
import shutil
import unittest
from PIL import Image

import pptx
from pptx.util import Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from merger_engine import PPTXMerger, PDFExporter, validate_pptx_file


class TestPPTUnify(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.test_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_scratch")
        os.makedirs(cls.test_dir, exist_ok=True)

    @classmethod
    def tearDownClass(cls):
        if os.path.exists(cls.test_dir):
            shutil.rmtree(cls.test_dir)

    def _create_sample_deck(self, filename: str, slide_titles: list, include_image: bool = False) -> str:
        prs = pptx.Presentation()
        for idx, title_text in enumerate(slide_titles):
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            slide.shapes.title.text = title_text
            slide.placeholders[1].text = f"Slide content for {title_text}"

            if include_image and idx == 0:
                # Add a test PNG
                img = Image.new("RGB", (80, 80), color=(50, 120, 220))
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                buf.seek(0)
                slide.shapes.add_picture(buf, Inches(1), Inches(4), Inches(2), Inches(2))

        path = os.path.join(self.test_dir, filename)
        prs.save(path)
        return path

    def test_01_validation(self):
        """Test validation on valid, empty, and non-pptx files."""
        deck = self._create_sample_deck("valid.pptx", ["Slide 1", "Slide 2"])
        val = validate_pptx_file(deck)
        self.assertTrue(val["is_valid"])
        self.assertEqual(val["slide_count"], 2)

        # Empty file
        empty_path = os.path.join(self.test_dir, "empty.pptx")
        with open(empty_path, "wb") as f:
            pass
        val_empty = validate_pptx_file(empty_path)
        self.assertFalse(val_empty["is_valid"])
        self.assertIn("empty", val_empty["error_message"].lower())

        # Non-pptx file
        bad_path = os.path.join(self.test_dir, "corrupt.pptx")
        with open(bad_path, "wb") as f:
            f.write(b"not a valid zip or pptx")
        val_bad = validate_pptx_file(bad_path)
        self.assertFalse(val_bad["is_valid"])

    def test_02_sequential_merge(self):
        """Test sequential line-by-line merge of 3 decks."""
        deck1 = self._create_sample_deck("seq1.pptx", ["Deck1 S1", "Deck1 S2"], include_image=True)
        deck2 = self._create_sample_deck("seq2.pptx", ["Deck2 S1"])
        deck3 = self._create_sample_deck("seq3.pptx", ["Deck3 S1", "Deck3 S2", "Deck3 S3"])

        out_path = os.path.join(self.test_dir, "merged_seq.pptx")
        merger = PPTXMerger([deck1, deck2, deck3])
        result = merger.merge(out_path, mode="sequential")

        self.assertTrue(result["success"])
        self.assertEqual(result["total_slides"], 6)
        self.assertTrue(os.path.exists(out_path))

        # Check merged presentation
        prs = pptx.Presentation(out_path)
        self.assertEqual(len(prs.slides), 6)

    def test_03_interleaved_merge(self):
        """Test interleaved line-by-line merge."""
        deck1 = self._create_sample_deck("intl1.pptx", ["A1", "A2"])
        deck2 = self._create_sample_deck("intl2.pptx", ["B1", "B2"])

        out_path = os.path.join(self.test_dir, "merged_intl.pptx")
        merger = PPTXMerger([deck1, deck2])
        result = merger.merge(out_path, mode="interleaved")

        self.assertTrue(result["success"])
        self.assertEqual(result["total_slides"], 4)
        self.assertTrue(os.path.exists(out_path))

    def test_04_pdf_export(self):
        """Test PDF export engine."""
        deck = self._create_sample_deck("for_pdf.pptx", ["PDF Slide 1", "PDF Slide 2"], include_image=True)
        pdf_out = os.path.join(self.test_dir, "output.pdf")
        success = PDFExporter.convert_to_pdf(deck, pdf_out)
        self.assertTrue(success, "PDF export should succeed via Keynote or ReportLab fallback")
        self.assertTrue(os.path.exists(pdf_out))
        self.assertGreater(os.path.getsize(pdf_out), 0)

    def test_05_scale_100_files(self):
        """Stress test: Upload and merge 100 PPTX presentations error-free."""
        print("\n--> Generating 100 test PPTX presentations...")
        decks = []
        for i in range(1, 101):
            deck_path = self._create_sample_deck(f"batch_deck_{i:03d}.pptx", [f"Deck {i} Slide 1"])
            decks.append(deck_path)

        self.assertEqual(len(decks), 100)

        # Validate all 100
        print("--> Validating all 100 presentations...")
        for p in decks:
            v = validate_pptx_file(p)
            self.assertTrue(v["is_valid"])
            self.assertEqual(v["slide_count"], 1)

        # Merge all 100
        print("--> Merging all 100 presentations into one PPTX...")
        out_100 = os.path.join(self.test_dir, "merged_100_decks.pptx")
        merger = PPTXMerger(decks)
        result = merger.merge(out_100, mode="sequential")

        self.assertTrue(result["success"])
        self.assertEqual(result["total_slides"], 100)
        self.assertEqual(result["merged_files_count"], 100)
        self.assertTrue(os.path.exists(out_100))
        print(f"--> [PASS] Successfully merged 100 presentations into {out_100} ({os.path.getsize(out_100) / 1024:.1f} KB)")


if __name__ == "__main__":
    unittest.main()
