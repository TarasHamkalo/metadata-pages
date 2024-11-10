import logging
from importlib.metadata import metadata

from metadata import read_metadata_from_docx, read_metadata_from_doc
from exiftool import SimpleExifTool


def write_metadata_to_html(file_paths: list[str], output_html: str):
  with open(output_html, 'w', encoding='utf-8') as html_file:
    # Write the HTML header
    html_file.write(
      '<html><head><meta http-equiv="Content-Type" content="text/html;charset=UTF-8"/><title>Metadata Report</title></head><body>\n')
    html_file.write('<h1>Metadata Report</h1>\n')
    html_file.write('<table border="1">\n')

    # Header row
    headers = [
      'Path', 'Creator', 'Last Modified By', 'Date Created', 'Date Modified',
      'Last Printed', 'Template', 'Total Time', 'Pages'
    ]
    html_file.write('<tr>' + ''.join(f'<th>{header}</th>' for header in headers) + '</tr>\n')

    # Process each file and append metadata rows
    for file_path in file_paths:
      try:
        metadata = read_metadata_from_docx(file_path)
        row_data = [
          metadata.path,
          metadata.creator or '',
          metadata.lastModifiedBy or '',
          metadata.DateCreated or '',
          metadata.DateModified or '',
          metadata.lastPrinted or '',
          metadata.template or '',
          metadata.totalTime or '',
          metadata.pages or ''
        ]

        # Write a table row with the metadata
        html_file.write('<tr>' + ''.join(f'<td>{data}</td>' for data in row_data) + '</tr>\n')
      except Exception as e:
        logging.error(f'Error processing {file_path}: {e}')

    # Close the HTML tags
    html_file.write('</table>\n')
    html_file.write('</body></html>\n')

  print(f'Metadata written to {output_html}')


if __name__ == "__main__":
  allDocInfo = []
  # getAllFiles(allDocInfo, '/home/taras-hamkalo/other/metadata-pages/19-20/processed/docx')
  # savetoCSV(allDocInfo, 'results.csv')
  # getFileInfo('/home/taras-hamkalo/other/metadata-pages/19-20/processed/docx/4064284486.docx')
  # get_file_info('/home/taras-hamkalo/other/metadata-pages/19-20/processed/docx/4064284486.docx')
  path = 'demo.doc'
  # write_metadata_to_html([path], 'report.html')
  # path = 'demo.docx'
  # metadata = read_metadata_from_docx(path)
  # import exiftool
  # with exiftool.ExifToolHelper(common_args=['-G1', '-n']) as et:
  #   metadata = read_metadata_from_doc(path, et)
  import exiftool
  with SimpleExifTool() as m:
    metadata = read_metadata_from_doc(path, m)
  print(metadata)

  print('DONE')
