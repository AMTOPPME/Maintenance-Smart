import json
import pandas as pd
from pathlib import Path

# === 設定區 ===
# 1. 將 Download 下來的 maintenance_logs.json 放在跟這支程式同一個資料夾
JSON_FILE = Path("maintenance_logs.json")
# 2. 輸出的 Excel 報表名稱
OUTPUT_FILE = Path("Maintenance_Report_From_JSON.xlsx")


def load_logs(json_path: Path) -> pd.DataFrame:
    """Load maintenance logs from JSON file into a standardized DataFrame."""
    if not json_path.exists():
        raise FileNotFoundError(f"JSON file not found: {json_path}")

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, list):
        raise ValueError("JSON format error: root should be a list of records")

    df = pd.DataFrame(data)

    # ---- 日期欄位：支援 Date / date，統一成 Date ----
    if "Date" in df.columns:
        date_col = "Date"
    elif "date" in df.columns:
        date_col = "date"
    else:
        raise ValueError("Missing 'Date' or 'date' field in JSON records")

    df["Date"] = pd.to_datetime(df[date_col])
    df["month"] = df["Date"].dt.to_period("M").astype(str)

    # ---- Downtime 欄位：支援 downtime_min / downtime，統一成 downtime_min ----
    if "downtime_min" in df.columns:
        dt_col = "downtime_min"
    elif "downtime" in df.columns:
        dt_col = "downtime"
    else:
        dt_col = None

    if dt_col:
        df["downtime_min"] = pd.to_numeric(df[dt_col], errors="coerce").fillna(0)
    else:
        df["downtime_min"] = 0

    # ---- 文字欄位：有新欄位就用新欄位，沒有就 fallback 到舊欄位 ----
    def pick(col_new: str, col_old: str | None = None) -> pd.Series:
        if col_new in df.columns:
            s = df[col_new]
        elif col_old and col_old in df.columns:
            s = df[col_old]
        else:
            s = ""
        return s.fillna("")

    df["line"] = pick("line", "line")
    df["section"] = pick("section", "section")
    df["asset_id"] = pick("asset_id", "equipment")         # 新 asset_id，舊 equipment
    df["category"] = pick("category", "category")
    df["root_cause"] = pick("root_cause", "rootcause")
    df["action_taken"] = pick("action_taken", "action")
    df["location_from"] = pick("location_from")
    df["location_to"] = pick("location_to")
    df["timestamp"] = pick("timestamp")
    df["id"] = pick("id")

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
        df.groupby("line")["downtime_min"]
        .sum()
        .reset_index(name="TotalDowntimeMin")
        .sort_values("TotalDowntimeMin", ascending=False)
    )
    summary["02_ByLine_DowntimeMin"] = by_line_dt

    # 3. Monthly downtime
    by_month_dt = (
        df.groupby("month")["downtime_min"]
        .sum()
        .reset_index(name="TotalDowntimeMin")
        .sort_values("month")
    )
    summary["03_ByMonth_DowntimeMin"] = by_month_dt

    # 4. Category stats
    by_cat = (
        df.groupby("category")["downtime_min"]
        .agg(
            Count="size",
            TotalDowntimeMin="sum",
            AvgDowntimeMin="mean"
        )
        .reset_index()
        .sort_values("TotalDowntimeMin", ascending=False)
    )
    summary["04_ByCategory"] = by_cat

    # 5. Top 20 Asset IDs by count
    by_asset_count = (
        df.groupby("asset_id")
        .size()
        .reset_index(name="Count")
        .sort_values("Count", ascending=False)
        .head(20)
    )
    summary["05_TopAssetID_Count"] = by_asset_count

    # 6. Top 20 Asset IDs by downtime
    by_asset_dt = (
        df.groupby("asset_id")["downtime_min"]
        .sum()
        .reset_index(name="TotalDowntimeMin")
        .sort_values("TotalDowntimeMin", ascending=False)
        .head(20)
    )
    summary["06_TopAssetID_DowntimeMin"] = by_asset_dt

    # 7. Raw data (sorted) － 用比較好看的欄位名稱輸出
    raw_cols = [
        "Date",
        "line",
        "section",
        "asset_id",
        "category",
        "downtime_min",
        "root_cause",
        "action_taken",
        "location_from",
        "location_to",
        "timestamp",
        "id",
    ]
    raw = df[raw_cols].sort_values("Date").rename(
        columns={
            "line": "Line",
            "section": "Section",
            "asset_id": "Asset ID",
            "category": "Category",
            "downtime_min": "Downtime (min)",
            "root_cause": "Root cause",
            "action_taken": "Action taken",
            "location_from": "Location from",
            "location_to": "Location to",
            "timestamp": "Timestamp",
            "id": "ID",
        }
    )
    summary["99_Raw_Logs"] = raw

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
