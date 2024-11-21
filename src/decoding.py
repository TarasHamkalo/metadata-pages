from src.constants import VALID_CHARACTERS_REGEX, SOURCE_ENCODINGS
import unicodedata

def validate_decoded_filename(s: str) -> bool:
    return s and s.isprintable() and VALID_CHARACTERS_REGEX.match(s)


def decode_from_eu_central(data: bytes) -> str | None:
    for encoding in SOURCE_ENCODINGS:
        try:
            decoded = data.decode(encoding)
            normdecoded = unicodedata.normalize('NFC', decoded)
            if validate_decoded_filename(normdecoded):
                return normdecoded

        except UnicodeDecodeError as ignore:
            continue
    return None

def decode_nullable(data: bytes) -> str | None:
    if data:
        return decode_from_eu_central(data)
    return None

def decode_from_cp437(s: str) -> str | None:
    try:
        return decode_from_eu_central(s.encode('cp437'))
    except:
        return None