import logging
import os
import xml.dom.minidom
import zipfile
from datetime import datetime, date
from pathlib import Path
from typing import List, Dict

from olefile import OleMetadata, olefile

from .simple_exiftool import SimpleExifTool
from src.decoding import decode_nullable

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

    for docx_path in filetype_to_paths['docx']:
        try:
            metadatas.append(read_metadata_from_docx(docx_path))
        except Exception as e:
            logging.warning(f"Error extracting metadata from {docx_path}.\nCause: {e}")

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

        # date_format = "%Y-%m-%d %H:%M:%s"  # "2021-12-09 20:08:00"
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
