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

    # Basic column checks & conversions
    if "date" not in df.columns:
        raise ValueError("Missing 'date' field in JSON records")

    df["date"] = pd.to_datetime(df["date"])
    df["month"] = df["date"].dt.to_period("M").astype(str)

    if "downtime" in df.columns:
        df["downtime"] = pd.to_numeric(df["downtime"], errors="coerce").fillna(0)
    else:
        df["downtime"] = 0

    # Fill missing text fields with empty string to avoid NaN
    for col in ["line", "section", "equipment", "category", "rootcause", "action"]:
        if col not in df.columns:
            df[col] = ""
        else:
            df[col] = df[col].fillna("")

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
