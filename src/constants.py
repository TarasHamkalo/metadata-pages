import re

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

TABLE_HEADERS = [
    'Filename', 'Filetype', 'Submitter', 'Creator', 'Last Modified By', 'Total edit time (MIN)',
    'Date Created', 'Date Modified', 'Last Printed', 'Template', 'Pages'
]

VALID_CHARACTERS = "\/01234567789()áäčďéíĺľňóôŕšťúýžabcdefghijklmnopqrstuvwxyz@#$&. ,-_*"
VALID_CHARACTERS_REGEX = re.compile(f"^[{re.escape(VALID_CHARACTERS)}]+$", re.IGNORECASE)

SOURCE_ENCODINGS = ['cp852', 'utf-8', 'cp1252', 'latin2']
