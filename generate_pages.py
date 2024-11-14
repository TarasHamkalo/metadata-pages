import argparse
import csv
import logging
import os
import re
import sys
import tempfile
import xml
import xml.dom.minidom
import zipfile
from datetime import datetime, date
from pathlib import Path
from typing import List, Dict
from zipfile import ZipFile

import olefile
from olefile import OleMetadata

from simple_exiftool import SimpleExifTool

logging.getLogger().setLevel(logging.DEBUG)

TABLE_HEADERS = [
  'Filename', 'Filetype', 'Submitter', 'Creator', 'Last Modified By', 'Total edit time (MIN)',
  'Date Created', 'Date Modified', 'Last Printed', 'Template', 'Pages'
]


def get_row_data(metadata, submitter) -> List[str]:
  return [
    metadata.filename,
    metadata.extension or '',
    submitter,
    metadata.creator or '',
    metadata.last_modified_by or '',
    metadata.total_time or '0',
    metadata.date_created.date() if metadata.date_created else '',
    metadata.date_modified.date() if metadata.date_modified else '',
    metadata.last_printed.date() if metadata.last_printed else '',
    metadata.template or '',
    metadata.pages or ''
  ]


HTML_TABLE_STYLES = """
        <style>
            table {
                width: 100%;
                border-collapse: collapse;
                font-family: Arial, sans-serif;
                font-size: 14px;
            }

            th, td {
                padding: 10px;
                text-align: left;
                border-bottom: 1px solid #ddd;
            }

            th {
                background-color: #f2f2f2;
                font-weight: bold;
            }

            tr:nth-child(even) {
                background-color: #f9f9f9;
            }

            tr:hover {
                background-color: #e0e0e0;
            }

            caption {
                font-size: 18px;
                font-weight: bold;
                margin-bottom: 10px;
            }
        </style>
    """


class Metadata:

  def __init__(self, path):
    self.path = path
    self.filename = os.path.basename(self.path)
    _, self.extension = os.path.splitext(self.path)
    self.extension = self.extension[1:]

    self.pages: str | None = None
    self.template: str | None = None

    self.total_time: int | None = None  # In minutes

    self.creator: str | None = None
    self.last_modified_by: str | None = None

    self.date_created: datetime | None = None
    self.date_modified: datetime | None = None
    self.last_printed: datetime | None = None


def read_metadata_recursively(path: Path) -> List[Metadata]:
  if not path.is_dir():
    logging.warning(f"Path is not a directory: {path}")
    return []

  filetype_to_paths = collect_metadata_paths(path)

  metadatas: List[Metadata] = []
  for doc_path in filetype_to_paths['doc']:  # type: Path
    metadatas.append(read_metadata_from_doc(doc_path))

  if len(filetype_to_paths['pdf']) > 0:
    try:
      with SimpleExifTool() as exif_tool:
        for pdf_path in filetype_to_paths['pdf']:  # type: Path
          metadatas.append(read_metadata_from_pdf(pdf_path, exif_tool))
    except Exception as e:
      logging.error(f"Error extracting metadata from doc or pdf format, "
                    f"perhaps Exiftool is not installed.\n"
                    f"Cause: {e}"
      )

  for docx_path in filetype_to_paths['docx']:
    try:
      metadatas.append(read_metadata_from_docx(docx_path))
    except Exception as e:
      logging.warning(f"Error extracting metadata from {docx_path}.\nCause: {e}")

  return metadatas


def collect_metadata_paths(path) -> Dict[str, List[Path]]:
  filetype_to_paths: Dict[str, List[Path]] = {
    'docx': [],
    'doc': [],
    'pdf': []
  }

  for child in Path(path).rglob('*'):  # type: Path
    if child.is_dir():
      continue

    if child.name.startswith('.'):
      # logging.warning(f"Found hidden file {path}")
      continue

    filetype = child.suffix.lower()[1:]
    if filetype in filetype_to_paths.keys():
      logging.info(f"Found submission file {child}")
      filetype_to_paths[filetype].append(child)
  return filetype_to_paths


def read_metadata_from_docx(path: Path) -> Metadata:
  metadata: Metadata = Metadata(path)
  date_format = "%Y-%m-%dT%H:%M:%SZ"  # 2021-12-20T18:41:00Z
  with zipfile.ZipFile(str(path), 'r') as zipf:
    try:
      core = xml.dom.minidom.parseString(zipf.read('docProps/core.xml'))
      metadata.creator = get_dom_element_as_text(core, 'dc:creator')
      metadata.last_modified_by = get_dom_element_as_text(core, 'cp:lastModifiedBy')

      created = get_dom_element_as_text(core, 'dcterms:created')
      metadata.date_created = nullable_str_to_datetime(created, date_format)

      modified = get_dom_element_as_text(core, 'dcterms:modified')
      metadata.date_modified = nullable_str_to_datetime(modified, date_format)

      last_printed = get_dom_element_as_text(core, 'cp:lastPrinted')
      metadata.last_printed = nullable_str_to_datetime(last_printed, date_format)

    except Exception as e:
      logging.warning(f"Document does not have core xml: {path}")

    try:
      app = xml.dom.minidom.parseString(zipf.read('docProps/app.xml'))
      metadata.template = get_dom_element_as_text(app, 'Template')
      totalTime: str = get_dom_element_as_text(app, 'TotalTime')
      metadata.total_time = int(totalTime) if len(totalTime) > 0 else 0

      metadata.pages = get_dom_element_as_text(app, 'Pages')

    except Exception as e:
      logging.warning(f"Document does not have app xml: {path}")

  return metadata


def get_dom_element_as_text(doc, tag_name) -> str | None:
  try:
    return doc.getElementsByTagName(tag_name)[0].childNodes[0].data
  except (IndexError, AttributeError):
    return None


def decode_nullable(data: bytes) -> str | None:
  if data:
    return data.decode('iso-8859-2', errors='replace')
    # return data.decode('utf8', errors='replace')
  return None


def nullable_str_to_datetime(date: str | None, time_pattern: str) -> datetime | None:
  if date and len(date) > 0:
    return datetime.strptime(date, time_pattern)
  return None


def read_metadata_from_doc(path: Path) -> Metadata:
  metadata = Metadata(path)
  try:
    if not olefile.isOleFile(str(path)):
      logging.warning(f"Path is not a valid DOC file: {path}")
      return metadata

    date_format = "%Y-%m-%d %H:%M:%s"  # "2021-12-09 20:08:00"
    with olefile.OleFileIO(str(path)) as ofile:

      olemetadata: OleMetadata = ofile.get_metadata()
      metadata.total_time = olemetadata.total_edit_time
      metadata.template = decode_nullable(olemetadata.template)
      metadata.creator = decode_nullable(olemetadata.author)
      metadata.last_modified_by = decode_nullable(olemetadata.last_saved_by)
      metadata.date_created = olemetadata.create_time

      metadata.date_modified = olemetadata.last_saved_time
      metadata.last_printed = olemetadata.last_printed

      metadata.pages = olemetadata.num_pages
      # TODO: move this filtering to results
      if metadata.last_printed and metadata.last_printed.date() < date(1900, 1, 1):
        metadata.last_printed = None

      if metadata.total_time:
        metadata.total_time //= 60

  except Exception as e:
    logging.error(f"Error reading metadata for {path}: {e}")

  return metadata


def read_metadata_from_pdf(path: Path, exif_tool: SimpleExifTool) -> Metadata:
  metadata = Metadata(path)
  date_format = '%Y:%m:%d %H:%M:%S%z'  # 2021:12:14 17:52:05+00:00
  # modify_date_format = '%Y:%m:%d %H:%M:%S%z' # 2021:12:14 17:59:55Z
  try:
    exif_data = exif_tool.get_metadata(str(path))[0]
    metadata.pages = exif_data.get('PDF:PageCount')
    metadata.creator = exif_data.get('PDF:Creator')

    created = exif_data.get('PDF:CreateDate')
    metadata.date_created = nullable_str_to_datetime(created, date_format)
    modified = exif_data.get('PDF:ModifyDate')
    metadata.date_modified = nullable_str_to_datetime(modified, date_format)
  except Exception as e:
    logging.error(f"Error reading metadata for {path}: {e}")

  return metadata


def collect_from_zipped(path: Path) -> List[Metadata]:
  if not zipfile.is_zipfile(path):
    logging.warning(f"Not a zip file: {path}")
    return []

  with tempfile.TemporaryDirectory() as tempdir:
    # TODO: just try few
    with zipfile.ZipFile(path, 'r', metadata_encoding='iso-8859-2') as zf:  # type: ZipFile
      # for info in zf.infolist():
      #   filename = info.filename.encode('IBM852').decode('iso-8859-2')
      #   print(filename)
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


def write_metadata_to_html(dir_to_metadatas: Dict[Path, List[Metadata]], output_html: Path):
  with open(output_html, 'w', encoding='utf-8') as html_file:
    html_file.write(
      f"<html lang=sk>"
      f"<head>"
      f"""<meta http-equiv="Content-Type" content="text/html;charset=utf-8"/>"""
      f"{HTML_TABLE_STYLES}"
      f"<title>Metadata Report</title> </head> <body>"
    )

    html_file.write('<h1>Metadata Report</h1>\n')
    html_file.write('<table>\n')

    html_file.write('<tr>' + ''.join(f'<th>{header}</th>' for header in TABLE_HEADERS) + '</tr>\n')

    submitter_regex = re.compile(r"\d{4}_\d{4}_([A-Z][a-z]+_[A-Z][a-z]+)_")

    for directory, metadatas in dir_to_metadatas.items():  # type: Path, List[Metadata]
      submitter = extract_submitter(directory, submitter_regex)
      if len(metadatas) == 0:
        row_data = ['', '', submitter, '', '', '', '', '', '', '', '']
        html_file.write('<tr>' + ''.join(f'<td>{data}</td>' for data in row_data) + '</tr>\n')

      for metadata in metadatas:
        row_data = get_row_data(metadata, submitter)
        html_file.write('<tr>' + ''.join(f'<td>{data}</td>' for data in row_data) + '</tr>\n')

    html_file.write('</table>\n')
    html_file.write('</body></html>\n')

  print(f'Metadata written to {output_html}')


def extract_submitter(dir, submitter_regex):
  submitter_match = submitter_regex.search(str(dir))
  if submitter_match:
    submitter = submitter_match.group(1).replace('_', ' ')
  else:
    submitter = dir
  return submitter


def write_metadata_to_csv(dir_to_metadatas: Dict[Path, List[Metadata]], output_csv: Path):
  with open(output_csv, 'w', newline='', encoding='utf-8') as csv_file:
    writer = csv.writer(csv_file)
    writer.writerow(TABLE_HEADERS)
    submitter_regex = re.compile(r"\d{4}_\d{4}_([A-Z][a-z]+_[A-Z][a-z]+)_")
    for directory, metadatas in dir_to_metadatas.items():
      submitter = extract_submitter(str(directory), submitter_regex)
      if len(metadatas) == 0:
        row_data = ['', '', submitter, '', '', '', '', '', '', '', '']
        writer.writerow(row_data)

      for metadata in metadatas:
        row_data = get_row_data(metadata, submitter)
        writer.writerow(row_data)

  print(f'Metadata written to {output_csv}')


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

  # for info in zf.infolist():
  #   filename = ""
  #   try:
  #     filename = info.filename.encode('cp437').decode('utf-8', errors='ignore')
  #   except UnicodeDecodeError:
  #     filename = info.filename.encode('iso-8859-1').decode('utf-8', erors='ignore')
  #   print(filename)
  #
