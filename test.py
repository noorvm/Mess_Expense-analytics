import pandas as pd 

caution_details =pd.read_excel("caution details.xlsx")
Mess_caution =pd.read_excel("Mess caution.xlsx")

def clean_text(text):
    if pd.isna(text):
        return ""

    text = str(text).strip().lower()
    text = " ".join(text.split())

    return text

import re

def clean_mess_no(value):
    if pd.isna(value):
        return None

    value = str(value).strip().upper()

    match = re.search(r'(\d+)', value)

    if match:
        return int(match.group(1))

    return None
"""

print(Mess_caution.columns)
print("-----------------------------------------------------------------")
print(caution_details.columns)
"""
caution_details["Name_clean"] = caution_details["Name"].apply(clean_text)
Mess_caution["Name_clean"] = Mess_caution["Name"].apply(clean_text)

caution_details["MessNo_clean"] = caution_details["Mess No"].apply(clean_mess_no)
Mess_caution["MessNo_clean"] = Mess_caution["Mess No"].apply(clean_mess_no)

merged = pd.merge(
    Mess_caution,
    caution_details,
    on="MessNo_clean",
    how="left"
)

with pd.ExcelWriter("final_report.xlsx") as writer:
    merged.to_excel(writer, sheet_name="Merged Data", index=False)
    unmatched.to_excel(writer, sheet_name="Unmatched", index=False)