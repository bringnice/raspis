import json
import urllib.request
from datetime import datetime, timezone

from schedule_parser import parse_workbook

FILE_ID = "1RRrdDgLjqfFRYhjbxTdcJY_iNvP3ZEuu"
URL = f"https://drive.google.com/uc?export=download&id={FILE_ID}"

raw = urllib.request.urlopen(URL, timeout=60).read()
days = parse_workbook(raw)
if not days:
    raise SystemExit("Парсер ничего не нашёл — файл не скачался или изменилась структура")

with open("schedule.json", "w", encoding="utf-8") as f:
    json.dump({"days": days, "updated": datetime.now(timezone.utc).isoformat(timespec="minutes")},
              f, ensure_ascii=False, indent=1)
