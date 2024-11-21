import argparse
import logging
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import List, Dict
from zipfile import ZipFile

from src.decoding import decode_from_cp437
from src.reading.reading import Metadata, read_metadata_recursively
from src.report_writing import write_metadata_to_html, write_metadata_to_csv

logging.getLogger().setLevel(logging.DEBUG)

def collect_from_zipped(path: Path) -> List[Metadata]:
    if not zipfile.is_zipfile(path):
        logging.warning(f"Not a zip file: {path}")
        return []

    with tempfile.TemporaryDirectory() as tempdir:
        with zipfile.ZipFile(path, 'r') as zf:  # type: ZipFile
            for member in zf.infolist(): # type ZipInfo
                member.filename = decode_from_cp437(member.filename)

            zf.extractall(tempdir)
            tempdir_path = Path(tempdir)
            return read_metadata_recursively(tempdir_path)

def collect_metadata(input_dir: Path, zipped) -> Dict[Path, List[Metadata]]:
    dir_to_metadata = {}
    for subdir in input_dir.iterdir():
        metadatas = []
        if zipped:
            try:
                metadatas = collect_from_zipped(subdir)
            except Exception as e:
                logging.error(f"Was not able to extract metadata for {subdir}: \n{e}")
        else:
            try:
                metadatas = read_metadata_recursively(subdir)
            except Exception as e:
                logging.error(f"Was not able to extract metadata for {subdir}: \n{e}")
        logging.info(f"Dir {subdir} has {len(metadatas)} metadatas")

        dir_to_metadata[subdir] = metadatas

    return dir_to_metadata

def parse_args():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "input_dir",
        type=Path,
        help="Path to the input directory."
    )

    parser.add_argument(
        "--zipped",
        action="store_true",
        help="Specify this flag if directories inside input dir are zipped."
    )

    parser.add_argument(
        "--csv",
        action="store_true",
        help="Specify this flag if CSV output is required."
    )

    parser.add_argument(
        "--force",
        "-f",
        action="store_true",
        help="Overwrite existing output files if specified."
    )

    parser.add_argument(
        "--output-name",
        "-o",
        type=str,
        default="report",
        help="Name of the output file (without extension)."
    )

    return parser.parse_args()

def validate_output_files(html_path: Path, csv_path: Path, force: bool, csv_required: bool):
    if not force:
        if html_path.exists():
            print(f"HTML output file already exists: {html_path}")
            sys.exit(1)
        if csv_required and csv_path.exists():
            print(f"CSV output file already exists: {csv_path}")
            sys.exit(1)

def main():
    args = parse_args()

    html_output_path = Path(f"{args.output_name}.html")
    csv_output_path = Path(f"{args.output_name}.csv")

    input_dir = args.input_dir
    if not input_dir.exists() or not input_dir.is_dir():
        print(f"The path provided does not exist or is not a directory: {input_dir}")
        sys.exit(1)

    validate_output_files(html_output_path, csv_output_path, args.force, args.csv)

    dir_to_metadata = collect_metadata(input_dir, args.zipped)
    write_metadata_to_html(dir_to_metadata, html_output_path)
    if args.csv:
        write_metadata_to_csv(dir_to_metadata, csv_output_path)

if __name__ == "__main__":
    main()
