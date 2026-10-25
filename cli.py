"""
cli.py - Command Line Interface for PPT Unify
Allows merging up to 100 PPT/PPTX files and exporting to PPTX and PDF directly from terminal.

Usage:
    python cli.py file1.pptx file2.pptx -o merged.pptx --pdf
    python cli.py --dir /path/to/ppts/ -o output.pptx --mode interleaved --pdf
"""

import os
import sys
import argparse
import glob

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from merger_engine import PPTXMerger, PDFExporter, validate_pptx_file


def main():
    parser = argparse.ArgumentParser(description="PPT Unify - Merge up to 100 PowerPoint presentations and export to PPTX/PDF.")
    parser.add_argument("files", nargs="*", help="List of PPT/PPTX files to merge in order")
    parser.add_argument("-d", "--dir", help="Directory containing PPTX files to merge")
    parser.add_argument("-o", "--output", default="merged_presentation.pptx", help="Output PPTX path")
    parser.add_argument("-m", "--mode", choices=["sequential", "interleaved"], default="sequential", help="Merge strategy ('sequential' or 'interleaved')")
    parser.add_argument("--pdf", action="store_true", help="Also export merged result as PDF")
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose progress output")

    args = parser.parse_args()

    files = list(args.files)
    if args.dir:
        pattern = os.path.join(args.dir, "*.pptx")
        dir_files = sorted(glob.glob(pattern))
        files.extend(dir_files)

    if not files:
        parser.print_help()
        print("\n[!] Please specify files or a directory of PPTX files to merge.")
        sys.exit(1)

    print(f"\n[+] Loaded {len(files)} files to process.")

    # Validate each
    valid_files = []
    total_slides = 0
    for idx, f in enumerate(files):
        val = validate_pptx_file(f)
        if val["is_valid"]:
            valid_files.append(f)
            total_slides += val["slide_count"]
            if args.verbose:
                print(f"  [{idx+1}/{len(files)}] ✓ {os.path.basename(f)} ({val['slide_count']} slides, {val['aspect_ratio']})")
        else:
            print(f"  [{idx+1}/{len(files)}] ✗ {os.path.basename(f)}: {val['error_message']}")

    if not valid_files:
        print("\n[!] No valid presentations found to merge.")
        sys.exit(1)

    print(f"\n[+] Merging {len(valid_files)} valid presentations ({total_slides} total slides) with mode '{args.mode}'...")

    def progress(pct, msg):
        print(f"  [{pct:3d}%] {msg}")

    merger = PPTXMerger(valid_files)
    result = merger.merge(args.output, mode=args.mode, progress_callback=progress)

    print(f"\n[✓] Successfully merged to: {args.output}")
    print(f"    Total slides: {result['total_slides']}")
    print(f"    File size: {result['file_size'] / (1024*1024):.2f} MB")

    if args.pdf:
        pdf_path = os.path.splitext(args.output)[0] + ".pdf"
        print(f"\n[+] Exporting to PDF: {pdf_path}...")
        success = PDFExporter.convert_to_pdf(args.output, pdf_path, progress_callback=progress)
        if success:
            print(f"[✓] Successfully exported PDF to: {pdf_path}")
        else:
            print(f"[!] Warning: PDF export failed.")


if __name__ == "__main__":
    main()
