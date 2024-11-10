# import glob
# import xml.dom.minidom
# import zipfile
#
#
# def getAllFiles(allDocInfo, path):
#   for filename in glob.iglob(path + '/**/*.*', recursive=True):
#     allDocInfo.append(getFileInfo(filename))

import logging
import xml.dom.minidom
import zipfile
from importlib.metadata import metadata
from xml.dom.minidom import Document


class CoreProps:

  def __init__(self, dom: Document):
    self.creator: str = get_dom_element_as_text(dom, 'dc:creator')
    self.lastModifiedBy: str = get_dom_element_as_text(dom, 'cp:lastModifiedBy')
    self.DateCreated: str = get_dom_element_as_text(dom, 'dcterms:created')
    self.DateModified: str = get_dom_element_as_text(dom, 'dcterms:modified')
    self.lastPrinted: str = get_dom_element_as_text(dom, 'cp:lastPrinted')

class AppProps:

  def __init__(self, dom: Document):
    self.template: str = get_dom_element_as_text(dom, 'Template')
    self.totalTime: str = get_dom_element_as_text(dom, 'TotalTime')
    self.pages: str = get_dom_element_as_text(dom, 'Pages')

class Metadata:

  def __init__(self, app_props: AppProps, core_props: CoreProps):
    self.core_props: CoreProps = core_props
    self.app_props: AppProps = app_props

def read_metadata(file_path) -> Metadata:
  with zipfile.ZipFile(file_path, 'r') as f:
    print('Processing: ', file_path)
    core_props: CoreProps | None = None

    try:
      core_doc = xml.dom.minidom.parseString(f.read('docProps/core.xml'))
      core_props = CoreProps(core_doc)
    except Exception as e:
      logging.error('Document does not have core xml')

    try:
      app_doc = xml.dom.minidom.parseString(f.read('docProps/app.xml'))
      app_props = AppProps(app_doc)
    except Exception as e:
      logging.warning('Document does not have app xml')

    return Metadata(app_props, core_props)

def get_dom_element_as_text(doc, tag_name) -> str:
  try:
    return doc.getElementsByTagName(tag_name)[0].childNodes[0].data
  except (IndexError, AttributeError):
    return ''

def write_metadata_to_html(file_paths: list[str], output_html: str):
  # Create an HTML file with a table
  with open(output_html, 'w', encoding='utf-8') as html_file:
    # Write the HTML header
    html_file.write('<html><head><meta http-equiv="Content-Type" content="text/html;charset=UTF-8"/><title>Metadata Report</title></head><body>\n')

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
        metadata = read_metadata(file_path)
        core = metadata.core_props
        app = metadata.app_props

        row_data = [
          file_path,
          core.creator if core else '',
          core.lastModifiedBy if core else '',
          core.DateCreated if core else '',
          core.DateModified if core else '',
          core.lastPrinted if core else '',
          app.template if app else '',
          app.totalTime if app else '',
          app.pages if app else ''
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
  path = '../demo.docx'
  write_metadata_to_html([path], 'report.html')
  print('DONE')
