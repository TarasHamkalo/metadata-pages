#!/usr/bin/python3
import re
import zipfile
from pathlib import Path

# acceptable_chars = set('áäčďéíĺľňóôŕšťúýž') #ÁÄČĎÉÍĹĽŇÓÔŔŠŤÚÝŽ
# acceptable_chars = set('abcdefghijklmnopqrstuvwxyz')
# acceptable_chars = set('@#$&. ,-_')
VALID_CHARACTERS = "áäčďéíĺľňóôŕšťúýžabcdefghijklmnopqrstuvwxyz@#$&. ,-_"
VALID_FILENAME_REGEX = re.compile(f"^[{re.escape(VALID_CHARACTERS)}]+$", re.IGNORECASE)

def print_decoded_from_cp437(s: str, target_encoding: str):
  try: 
    print(f"Decoded using {target_encoding}: {s.encode('cp437').decode(target_encoding)}")
  except Exception as e:
    print(f"Failed to decode using {target_encoding}")

def decode_from_cp437(s: str, target_encoding: str) -> str | None:
  try: 
    return s.encode('cp437').decode(encoding=target_encoding, errors='strict')
  except:
    print(f"{target_encoding} error")
    return None

def extract_archive(path: Path) -> None:
  print(f"{path}")
  with zipfile.ZipFile(path, 'r') as z:
    for name in z.namelist():
      print(f"\t{name}")
      for encoding in ['cp852', 'utf-8', 'cp1252', 'latin2']:
        s = decode_from_cp437(name, encoding)
        if s and s.isprintable():
          print(f"\t\tDecoded using {encoding}: {s}")

def validate(s: str) -> bool:
  return s and s.isprintable() and VALID_FILENAME_REGEX.match(s)

def extract_with_name_validation(path: Path) -> None:
  print(f"{path}")
  with zipfile.ZipFile(path, 'r') as z:
    for name in z.namelist():
      print(f"\t{name}")
      for encoding in ['cp852', 'utf-8', 'cp1252', 'latin2']:
        s = decode_from_cp437(name, encoding)
        if validate(s):
          print(f"\t\tDecoded using {encoding}: {s}")

path = Path("/home/taras-hamkalo/Downloads/oct27/15-16")
# path = Path("/home/taras-hamkalo/Downloads/oct27/multi-file-doc")

for child in path.iterdir():
  if child.is_dir():
    continue

  extract_with_name_validation(child)
#

