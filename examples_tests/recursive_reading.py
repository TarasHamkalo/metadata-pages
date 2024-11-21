import tempfile
import zipfile
from typing import List
from zipfile import ZipFile

from loguru import logger

from src.reading.reading import read_metadata, Metadata


class Submission:

  def __init__(self, path):
    self.filename = path
    self.metadatas: List[Metadata] = []

  def __str__(self):
    return f"Sumbission=(path={self.filename}, metadata_count={len(self.metadatas)})"


def write_metadata_to_html(submissions: List[Submission], output_html: str):
  with open(output_html, 'w', encoding='ISO-8859-1') as html_file:
    # Write the HTML header
    html_file.write(
      '<html><head><meta http-equiv="Content-Type" content="text/html;charset=ISO-8859-1"/><title>Metadata Report</title></head><body>\n')
    html_file.write('<h1>Metadata Report</h1>\n')
    html_file.write('<table border="1">\n')

    # Header row
    headers = [
      'SubmissionPath', 'Filename', 'Filetype', 'Creator', 'Last Modified By', 'Date Created', 'Date Modified',
      'Last Printed', 'Template', 'Total Time', 'Pages'
    ]
    html_file.write('<tr>' + ''.join(f'<th>{header}</th>' for header in headers) + '</tr>\n')

    # Process each file and append metadata rows
    for submission in submissions:
      for metadata in submission.metadatas: # type: Metadata
        row_data = [
          submission.filename,
          metadata.filename,
          metadata.extension or '',
          metadata.creator or '',
          metadata.lastModifiedBy or '',
          metadata.dateCreated or '',
          metadata.dateModified or '',
          metadata.lastPrinted or '',
          metadata.template or '',
          metadata.totalTime or '',
          metadata.pages or ''
        ]

        html_file.write('<tr>' + ''.join(f'<td>{data}</td>' for data in row_data) + '</tr>\n')

    # Close the HTML tags
    html_file.write('</table>\n')
    html_file.write('</body></html>\n')

  print(f'Metadata written to {output_html}')

def read_metadata_recursively(path) -> List[Metadata]:
  metadata_paths: List[Path] = []
  for path in Path(path).rglob('*'): # type: Path
    if path.name.startswith('.'):
      logger.warning(f"Found hidden file {path}")
      continue

    if path.suffix.lower() in ['.docx', '.doc', '.pdf']:
      logger.info(f"Found submission file {path}")
      metadata_paths.append(path)

  return [read_metadata(str(path)) for path in metadata_paths]


if __name__ == "__main__":
  INPUT_DIR = "/19-20/test"
  # UNZIPPED_DIR = "/home/taras-hamkalo/repositories/metadata-pages/19-20/unzipped"
  # ZIPPED_DIR = "/home/taras-hamkalo/repositories/metadata-pages/19-20/zipped"
  # UNZIPPED_DIR = "/home/taras-hamkalo/other/data/ipc/meged-before-23-24/docx"
  is_zipped = False
  # is_zipped = True

  from pathlib import Path

  submissions = []
  for path in Path(INPUT_DIR).glob('*'): # type: Path
    submission = Submission(path.name)
    if is_zipped:
      with tempfile.TemporaryDirectory() as tempdir:
        with zipfile.ZipFile(str(path), 'r') as zip_ref: # type: ZipFile
          zip_ref.extractall(tempdir)
          submission.metadatas = read_metadata_recursively(tempdir)
    else:
      submission.metadatas = read_metadata_recursively(path)

    submissions.append(submission)


    # print(metadata)
  write_metadata_to_html(submissions, output_html='report.html')
  #
  # for submission in submissions:
  #   print(submission)
  #   for metadata in submission.metadatas:
  #     print(metadata)

  print('DONE')
