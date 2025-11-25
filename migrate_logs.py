import json
from pathlib import Path

INPUT_FILE = Path("maintenance_logs.json")
OUTPUT_FILE = Path("maintenance_logs_new.json")

def migrate():
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"Could not find {INPUT_FILE}")

    with open(INPUT_FILE, "r", encoding="utf-8") as f:
        old_data = json.load(f)

    if not isinstance(old_data, list):
        raise ValueError("JSON format error: root should be a list of records")

    new_data = []

    for r in old_data:
        new_record = {
            # 舊 date → 新 Date
            "Date": r.get("date", ""),
            "line": r.get("line", ""),
            "section": r.get("section", ""),

            # 舊 equipment → 新 asset_id
            "asset_id": r.get("equipment", ""),

            "category": r.get("category", ""),

            # 舊 downtime → 新 downtime_min
            "downtime_min": r.get("downtime", 0),

            # 舊 rootcause / action → 新 root_cause / action_taken
            "root_cause": r.get("rootcause", ""),
            "action_taken": r.get("action", ""),

            # 新增欄位（舊資料沒有 → 先給空字串，之後新表單會寫入真正內容）
            "location_from": "",
            "location_to": "",

            # 其他欄位保留
            "timestamp": r.get("timestamp", ""),
            "id": r.get("id", "")
        }

        new_data.append(new_record)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as out:
        json.dump(new_data, out, indent=2, ensure_ascii=False)

    print(f"Migration completed! New file created:\n{OUTPUT_FILE.resolve()}")


if __name__ == "__main__":
    migrate()
