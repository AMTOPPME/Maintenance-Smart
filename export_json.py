import json
import pandas as pd
from pathlib import Path

# 設定：你的 JSON 檔案
JSON_FILE = Path("ms-maintenance.json") 
OUTPUT_FILE = Path("Maintenance_Report_From_JSON.xlsx")

with open(JSON_FILE, "r", encoding="utf-8") as f:
    data = json.load(f)

df = pd.DataFrame(data)

# 轉型
df["date"] = pd.to_datetime(df["date"])
df["downtime"] = pd.to_numeric(df["downtime"], errors="coerce").fillna(0)
df["month"] = df["date"].dt.to_period("M").astype(str)

# 統計
summary = {
    "ByLine_Count": df.groupby("line").size().reset_index(name="Count"),
    "ByLine_Downtime": df.groupby("line")["downtime"].sum().reset_index(),
    "ByMonth": df.groupby("month")["downtime"].sum().reset_index(),
    "ByCategory": df.groupby("category")["downtime"].sum().reset_index(),
    "Raw_Data": df
}

# 輸出多個 sheet
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    for name, table in summary.items():
        table.to_excel(writer, sheet_name=name[:31], index=False)

print("Export done →", OUTPUT_FILE)
