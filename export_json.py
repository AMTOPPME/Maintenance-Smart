import json
import pandas as pd
from pathlib import Path

# === 設定區 ===
# 1. 將 Download 下來的 maintenance_logs.json 放在跟這支程式同一個資料夾
JSON_FILE = Path("maintenance_logs.json")
# 2. 輸出的 Excel 報表名稱
OUTPUT_FILE = Path("Maintenance_Report_From_JSON.xlsx")


def load_logs(json_path: Path) -> pd.DataFrame:
    """Load maintenance logs from JSON file into a DataFrame."""
    if not json_path.exists():
        raise FileNotFoundError(f"JSON file not found: {json_path}")

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, list):
        raise ValueError("JSON format error: root should be a list of records")

    df = pd.DataFrame(data)

    # ====== 日期欄位（支援 date / Date）======
    if "date" not in df.columns:
        if "Date" in df.columns:
            df["date"] = df["Date"]
        else:
            raise ValueError(
                "Missing 'date' field in JSON records (also can't find 'Date'). "
                f"Available columns: {list(df.columns)}"
            )

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["month"] = df["date"].dt.to_period("M").astype(str)

    # ====== Downtime 欄位（支援 downtime / downtime_min）======
    if "downtime" in df.columns:
        pass
    elif "downtime_min" in df.columns:
        df["downtime"] = df["downtime_min"]
    else:
        df["downtime"] = 0

    df["downtime"] = pd.to_numeric(df["downtime"], errors="coerce").fillna(0)

    # ====== 文字欄位對應 ======
    # 你前端現在存的是：
    #   asset_id, root_cause, action_taken
    # 老的 Python 報表想要：
    #   equipment, rootcause, action
    col_map = {
        "equipment": ["asset_id"],
        "rootcause": ["root_cause"],
        "action": ["action_taken"],
    }

    # line / section / category 直接用原本名稱
    base_text_cols = ["line", "section", "category"]

    for col in base_text_cols:
        if col not in df.columns:
            df[col] = ""
        else:
            df[col] = df[col].fillna("")

    # 做欄位 mapping
    for target_col, candidates in col_map.items():
        if target_col in df.columns:
            df[target_col] = df[target_col].fillna("")
            continue

        src_col = None
        for c in candidates:
            if c in df.columns:
                src_col = c
                break

        if src_col is not None:
            df[target_col] = df[src_col].fillna("")
        else:
            df[target_col] = ""

    return df


def build_summary(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Build multiple summary tables from the raw logs."""
    summary: dict[str, pd.DataFrame] = {}

    # 1. Count by Line
    by_line_count = (
        df.groupby("line")
        .size()
        .reset_index(name="Count")
        .sort_values("Count", ascending=False)
    )
    summary["01_ByLine_Count"] = by_line_count

    # 2. Downtime by Line
    by_line_dt = (
        df.groupby("line")["downtime"]
        .sum()
        .reset_index(name="TotalDowntimeMin")
        .sort_values("TotalDowntimeMin", ascending=False)
    )
    summary["02_ByLine_DowntimeMin"] = by_line_dt

    # 3. Monthly downtime
    by_month_dt = (
        df.groupby("month")["downtime"]
        .sum()
        .reset_index(name="TotalDowntimeMin")
        .sort_values("month")
    )
    summary["03_ByMonth_DowntimeMin"] = by_month_dt

    # 4. Category stats
    by_cat = (
        df.groupby("category")["downtime"]
        .agg(
            Count="size",
            TotalDowntimeMin="sum",
            AvgDowntimeMin="mean"
        )
        .reset_index()
        .sort_values("TotalDowntimeMin", ascending=False)
    )
    summary["04_ByCategory"] = by_cat

    # 5. Top 20 equipments by count
    by_equipment_count = (
        df.groupby("equipment")
        .size()
        .reset_index(name="Count")
        .sort_values("Count", ascending=False)
        .head(20)
    )
    summary["05_TopEquipment_Count"] = by_equipment_count

    # 6. Top 20 equipments by downtime
    by_equipment_dt = (
        df.groupby("equipment")["downtime"]
        .sum()
        .reset_index(name="TotalDowntimeMin")
        .sort_values("TotalDowntimeMin", ascending=False)
        .head(20)
    )
    summary["06_TopEquipment_DowntimeMin"] = by_equipment_dt

    # 7. Raw data (sorted)
    summary["99_Raw_Logs"] = df.sort_values("date")

    return summary


def export_to_excel(tables: dict[str, pd.DataFrame], output_path: Path) -> None:
    """Write all summary tables into one Excel workbook."""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, table in tables.items():
            safe_name = sheet_name[:31]  # Excel sheet name limit
            table.to_excel(writer, sheet_name=safe_name, index=False)

    print(f"Report exported to: {output_path.resolve()}")


def main():
    print(f"Loading JSON logs from: {JSON_FILE}")
    df = load_logs(JSON_FILE)
    print(f"Loaded {len(df)} records.")

    print("Building summary tables...")
    summary_tables = build_summary(df)

    print("Exporting Excel report...")
    export_to_excel(summary_tables, OUTPUT_FILE)
    print("Done ✅")


if __name__ == "__main__":
    main()
