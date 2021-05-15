import pandas as pd


inSheet = pd.read_excel('export.xlsx', keep_default_na=False)
outSheet = pd.read_excel('myData.xlsx', keep_default_na=False)


for j in range(len(inSheet['Ad ID'])):
    entry = str(inSheet['Ad ID'][j][2:])
    for i in range(len(outSheet['Ad ID'])):
        ent = str(outSheet['Ad ID'][i])
        if(entry == ent):
            outSheet['Body'][i] = inSheet['Body'][j]
            break

with pd.ExcelWriter('Bridal-Wadsworth-Ads-Lifetime\ \(4\).xlsx', engine='xlsxwriter') as writer:
    outSheet.to_excel(writer, sheet_name='Sheet1')
