import re
from datetime import datetime, timezone
import random

from workspace.agent_params.values import TRANSLIT_DICT, ALLOWED_CHARS


def name_to_translit(text):
    translit_text = "".join(
        TRANSLIT_DICT.get(char.lower(), char) for char in text
    )
    return re.sub(r"\s+", "", translit_text).lower()[:75]

def get_time_stump() -> str:
    now = datetime.now(timezone.utc).astimezone()
    date = now.strftime("%Y-%m-%d %H:%M:%S.%f")
    time_zone = now.strftime("%z")
    return f"{date} {time_zone[:3]}:{time_zone[3:]}"

def get_timezone_offset() -> str:
    now = datetime.now(timezone.utc).astimezone()
    offset_seconds = now.utcoffset().total_seconds()
    hours = int(offset_seconds // 3600)
    minutes = abs(int((offset_seconds % 3600) // 60))
    sign = "+" if hours >= 0 else "-"
    return f"{sign}{abs(hours):02d}:{minutes:02d}"

def text_random(count=44, allowed_chars=ALLOWED_CHARS):
    return ''.join(random.choice(allowed_chars) for _ in range(count))