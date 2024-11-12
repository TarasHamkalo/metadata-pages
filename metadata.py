import logging
import xml.dom.minidom
import zipfile
import os
from importlib.metadata import metadata

from simple_exiftool import SimpleExifTool

class Metadata:

  def __init__(self, path):
    self.path = path
    self.pages: str | None
    self.template: str | None = None

    self.totalTime: int | None = None # In minutes

    self.creator: str | None = None
    self.dateCreated: str | None = None

    self.lastModifiedBy: str | None = None
    self.dateModified: str | None = None

    self.lastPrinted: str | None = None

def read_metadata(file_path) -> Metadata:
  _, extension = os.path.splitext(file_path)
  if extension == '.docx':
    return read_metadata_from_docx(file_path)
  elif extension == '.doc':
    return read_metadata_from_doc(file_path)
  elif extension == '.pdf':
    return read_metadata_from_pdf(file_path)

  return Metadata(file_path)

def read_metadata_from_docx(file_path) -> Metadata:
  metadata: Metadata = Metadata(file_path)

  with zipfile.ZipFile(file_path, 'r') as zipf:
    logging.info('Processing %s', file_path)
    try:
      core = xml.dom.minidom.parseString(zipf.read('docProps/core.xml'))
      metadata.creator = get_dom_element_as_text(core, 'dc:creator')
      metadata.lastModifiedBy = get_dom_element_as_text(core, 'cp:lastModifiedBy')
      metadata.dateCreated = get_dom_element_as_text(core, 'dcterms:created')
      metadata.dateModified = get_dom_element_as_text(core, 'dcterms:modified')
      metadata.lastPrinted = get_dom_element_as_text(core, 'cp:lastPrinted')
    except Exception as e:
      logging.warning('Document does not have core xml')

    try:
      app = xml.dom.minidom.parseString(zipf.read('docProps/app.xml'))
      metadata.template = get_dom_element_as_text(app, 'Template')
      totalTime: str= get_dom_element_as_text(app, 'TotalTime')
      metadata.totalTime = int(totalTime) if len(totalTime) > 0 else 0

      metadata.pages = get_dom_element_as_text(app, 'Pages')

    except Exception as e:
      logging.warning('Document does not have app xml')

  return metadata

def get_dom_element_as_text(doc, tag_name) -> str:
  try:
    return doc.getElementsByTagName(tag_name)[0].childNodes[0].data
  except (IndexError, AttributeError):
    return ''

def read_metadata_from_doc(file_path: str) -> Metadata:
  metadata = Metadata(file_path)
  with SimpleExifTool() as exif_tool:
    try:
      exif_data = exif_tool.get_metadata(file_path)[0]
      metadata.template = exif_data.get('MS-DOC:Template') or exif_data.get('FlashPix:Template')
      metadata.totalTime = exif_data.get('MS-DOC:TotalEditTime') or exif_data.get('FlashPix:TotalEditTime')
      metadata.pages = exif_data.get('MS-DOC:Pages') or exif_data.get('FlashPix:Pages')
      metadata.creator = exif_data.get('MS-DOC:Author') or exif_data.get('FlashPix:Author')
      metadata.lastModifiedBy = exif_data.get('MS-DOC:LastModifiedBy') or exif_data.get('FlashPix:LastModifiedBy')
      metadata.dateCreated = exif_data.get('MS-DOC:CreateDate') or exif_data.get('FlashPix:CreateDate')
      metadata.dateModified = exif_data.get('MS-DOC:ModifyDate') or exif_data.get('FlashPix:ModifyDate')
      metadata.lastPrinted = exif_data.get('MS-DOC:LastPrinted') or exif_data.get('FlashPix:LastPrinted')

      if metadata.totalTime:
        metadata.totalTime //= 60

    except Exception as e:
      logging.error(f"Error reading metadata for {file_path}: {e}")

  return metadata

def read_metadata_from_pdf(file_path) -> Metadata:
  metadata = Metadata(file_path)
  with SimpleExifTool() as exif_tool:
    try:
      exif_data = exif_tool.get_metadata(file_path)[0]
      metadata.pages = exif_data.get('PDF:PageCount')
      metadata.creator = exif_data.get('PDF:Creator')
      metadata.dateCreated = exif_data.get('PDF:CreateDate')
      metadata.dateModified = exif_data.get('PDF:ModifyDate')
    except Exception as e:
      logging.error(f"Error reading metadata for {file_path}: {e}")

  return metadata

#   [PDF]           Producer                        : macOS Verzia 10.15.2 (Zostava 19C57) Quartz PDFContext
# [PDF]           Creator                         : Word
# [PDF]           Create Date                     : 2019:12:15 19:21:18Z
# [PDF]           Modify Date                     : 2019:12:15 19:21:18Z
# def read_metadata_from_files(file_paths: list[str]) -> list[Metadata]:
#   metadata_list = []
#   # Use a single ExifTool process for all files
#   with ExifToolHelper(common_args=['-G1', '-n']) as et:
#     for file_path in file_paths:
#       if file_path.lower().endswith(('.doc', '.docx')):
#         metadata = read_metadata_from_doc(file_path, et)
#         metadata_list.append(metadata)
#       else:
#         logging.warning(f"Unsupported file type: {file_path}")
#
#   return metadata_list