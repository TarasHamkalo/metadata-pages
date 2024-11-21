import csv
import re
from pathlib import Path
from typing import Dict, List

from src.constants import TABLE_HEADERS, HTML_TABLE_STYLES
from src.reading.reading import Metadata

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

def write_metadata_to_html(dir_to_metadatas: Dict[Path, List[Metadata]], output_html: Path):
    with open(output_html, 'w', encoding='utf-8') as html_file:
        html_file.write(
            f"<html lang=sk><head>"
            f"""<meta http-equiv="Content-Type" content="text/html;charset=utf-8"/>"""
            f"{HTML_TABLE_STYLES}"
            f"<title>Metadata Report</title> </head> <body>"
        )

        html_file.write('<h1>Metadata Report</h1>\n')

        table_headers = ''.join(f'<th>{header}</th>' for header in TABLE_HEADERS)
        html_file.write('<table><tr>' + table_headers + '</tr>\n')

        submitter_regex = re.compile(r"\d{4}_\d{4}_([A-Z][a-z]+_[A-Z][a-z]+)_")

        for directory, metadatas in dir_to_metadatas.items():  # type: Path, List[Metadata]
            submitter = extract_submitter(directory, submitter_regex)
            if len(metadatas) == 0:
                row_data = ['', '', submitter, '', '', '', '', '', '', '', '']
                html_file.write(
                    '<tr>' + ''.join(f'<td>{data}</td>' for data in row_data) + '</tr>\n')

            for metadata in metadatas:
                row_data = get_row_data(metadata, submitter)
                html_file.write(
                    '<tr>' + ''.join(f'<td>{data}</td>' for data in row_data) + '</tr>\n')

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

