"""
Парсер файлу Dodatky.md — витягує дані про населені пункти та КСП роти.
"""
import re
from datetime import datetime
from typing import List, Dict, Optional


def parse_dodatky(filepath: str) -> List[Dict]:
    """
    Парсить md-таблицю з Dodatky.md.

    Формат таблиці:
    |***Період перебування***|населений_пункт|КСП_РОТИ|
    |-|-|-|
    |***31.08.2025***|Тищенківка|ВЕЛИКІ ХУТОРИ|

    Returns:
        [{date: datetime, населений_пункт: str, КСП_РОТИ: str}, ...]
    """
    with open(filepath, "r", encoding="utf-8") as f:
        content = f.read()

    entries = []
    for line in content.splitlines():
        line = line.strip()
        if not line.startswith("|"):
            continue
        # Пропускаємо заголовок та розділювач
        if "Період" in line or line.replace("|", "").replace("-", "").strip() == "":
            continue

        parts = [p.strip() for p in line.split("|") if p.strip()]
        if len(parts) < 3:
            continue

        # Витягуємо дату з ***dd.mm.yyyy*** або просто dd.mm.yyyy
        date_str = re.sub(r"\*+", "", parts[0]).strip()
        # Також прибираємо зворотні слеші (markdown escape)
        date_str = date_str.replace("\\", "")
        try:
            dt = datetime.strptime(date_str, "%d.%m.%Y")
        except ValueError:
            continue

        location = parts[1].replace("\\", "").strip()
        ksp = parts[2].replace("\\", "").strip()

        entries.append({
            "date": dt,
            "населений_пункт": location,
            "КСП_РОТИ": ksp,
        })

    # Сортуємо за датою
    entries.sort(key=lambda e: e["date"])
    return entries


def get_location_for_date(filepath: str, br_date: datetime) -> Dict[str, str]:
    """
    Повертає {населений_пункт, КСП_РОТИ} для дати БР.
    Знаходить останній період, дата якого <= br_date.

    Returns:
        {населений_пункт: str, КСП_РОТИ: str} або {населений_пункт: "—", КСП_РОТИ: "—"}
    """
    import os
    if not os.path.exists(filepath):
        return {"населений_пункт": "—", "КСП_РОТИ": "—"}

    entries = parse_dodatky(filepath)
    if not entries:
        return {"населений_пункт": "—", "КСП_РОТИ": "—"}

    # Знаходимо останній запис, де date <= br_date
    result = None
    for entry in entries:
        if entry["date"].date() <= br_date.date():
            result = entry
        else:
            break

    if result:
        return {
            "населений_пункт": result["населений_пункт"],
            "КСП_РОТИ": result["КСП_РОТИ"],
        }

    return {"населений_пункт": "—", "КСП_РОТИ": "—"}
