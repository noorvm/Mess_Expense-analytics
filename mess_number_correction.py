import pandas as pd

a=pd.read_excel("caution details.xlsx",sheet_name="march bill")

a["Mess No"] = a["Mess No"].str.replace("SWR-", "", regex=False)


with pd.ExcelWriter(
    "caution details.xlsx",
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace"
) as writer:
    a.to_excel(writer, sheet_name="march bill corrected", index=False)

print("New sheet created successfully!");