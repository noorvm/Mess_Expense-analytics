import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


ad=pd.read_excel("caution details.xlsx",sheet_name="admission details");
mb=pd.read_excel("caution details.xlsx",sheet_name="march bill");


mb["Mess no"] = (
    mb["Mess no"]
    .fillna(0)              # Replace empty cells with 0
    .astype(str)            # Convert everything to string
    .str.replace("SWR-", "", regex=False)
    .str.replace("swr-", "", regex=False)
)

mb["Mess no"] = pd.to_numeric(mb["Mess no"], errors="coerce").fillna(0).astype(int)

print(mb["Mess no"].dtype);

ad["Mess no"] = (
    ad["Mess no"]             
    .astype(str)            # Convert everything to string
    .str.replace("SWR-", "", regex=False)
)

ad["Mess no"] = pd.to_numeric(ad["Mess no"], errors="coerce").fillna(0).astype(int)

print(mb["Mess no"].dtype);

"""for i in range (len(ad)):
    for j in range(len(mb)):
        mb.loc[j,"Mess No"] == ad.loc[i,"Mess no"];
        new=pd.merge(ad,mb,on="Mess no",how ="inner")
"""
new=pd.merge(ad,mb,on="Mess no",how="outer");

""" new.to_excel("merged.xlsx", index=False);   """

print(new.columns.tolist());

new.drop(["Email","Unnamed: 0","MessCuts","Total MessCut Days","effective messcut","total working day","effective day","PER DAY","ESTABLISHMENT","FEAST","Unnamed: 14"], axis=1, inplace=True);
print("------------------------------------------------------------------------------\n");
print("------------------------------------------------------------------------------\n");

print(new.columns.tolist());

year1=new[(new["sem"]=="S1") | (new["sem"]=="S2")];
year2=new[(new["sem"]=="S3") | (new["sem"]=="S4")];
year3=new[(new["sem"]=="S5") | (new["sem"]=="S6")];
year4=new[(new["sem"]=="S7") | (new["sem"]=="S8")];

print(new["sem"].unique())

with pd.ExcelWriter("semester_tables.xlsx", engine="openpyxl") as writer:
    year1.to_excel(writer, sheet_name="Semester", startrow=0, index=False)
    year2.to_excel(writer, sheet_name="Semester", startrow=len(year1)+3, index=False)
    year3.to_excel(writer, sheet_name="Semester", startrow=len(year1)+len(year2)+6, index=False)
    year4.to_excel(writer, sheet_name="Semester", startrow=len(year1)+len(year2)+len(year3)+9, index=False) 








