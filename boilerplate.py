import pandas as pd
import calendar

df = pd.read_excel('C:\\GIT\\Ampol\\Data\\input1.xlsx', sheet_name='Sheet1', engine="openpyxl", header=None)

rowlength = len(df)
columnslength = len(df.columns)
dfhead = df.iloc[0].to_frame(name='col_0')
dfhead = dfhead.drop(dfhead[dfhead.col_0 == 'Units'].index)
dfhead = dfhead.dropna()
dfhead['year'] = ""
dfhead['period_type'] = ""
dfhead['period'] = ""
monthnames = ['Jan', 'Feb', 'Mar']
from datetime import datetime

for date in range(0, len(dfhead)):

    try:
        dfhead.iloc[date, 1] = dfhead.iloc[date, 0].year
        dfhead.iloc[date, 2] = 'Monthly'
        dfhead.iloc[date, 3] = calendar.month_name[dfhead.iloc[date, 0].month]
    except:

        dfhead.iloc[date, 1] = '20' + dfhead.iloc[date, 0].split("-")[1]
        dfhead.iloc[date, 2] = 'Quarterly'
        dfhead.iloc[date, 3] = dfhead.iloc[date, 0].split("-")[0]


product_name = ''
product_hier_1 = ''
product_hier_2 = ''
product_hier_3 = ''
column_names = ["product_name", "product_hier_1", "product_hier_2", "product_hier_3", "year", "period_type", "period",
                "unit", "value"]

finaldf = pd.DataFrame(columns=column_names)
csvrow = ''
for i in range(0, rowlength):
    for j in range(0, columnslength):
        if (df.iloc[i, j] == 'Product'):
            #print(df.iloc[i, j])
            product_name = df.iloc[i, j + 1]  # crude
            product_hier_1 = df.iloc[i + 1, j]  # physical
            product_hier_2 = df.iloc[i + 1, j + 1]  # physical premium
            for hierrow in range(i + 2, rowlength):
                if (df.iloc[hierrow, 0] == 'Product'):
                    break
                product_hier_3 = df.iloc[hierrow, j + 1]
                unit = df.iloc[hierrow, j + 2]
                dfhead['product_name'] = product_name
                dfhead['product_hier_1'] = product_hier_1
                dfhead['product_hier_2'] = product_hier_2
                dfhead['product_hier_3'] = product_hier_3

                df2 = df.loc[df[1] == product_hier_3].drop(columns=[0, 1, 2])
                df2 = df2.transpose()
                dfhead['value'] = df2
                dfhead['unit'] = unit

                dfhead = dfhead.reindex(columns=column_names)
                finaldf = pd.concat([finaldf, dfhead])

                dfhead.pop("product_name")  # ,"product_hier_1","product_hier_2","product_hier_3","unit","value")
                dfhead.pop("product_hier_1")
                dfhead.pop("product_hier_2")
                dfhead.pop("product_hier_3")
                dfhead.pop("unit")
                dfhead.pop("value")
                # csvrow=product_name+","+product_hier_1+","+product_hier_2+","+product_hier_3
                # print(csvrow)

finaldf.to_csv('C:\\GIT\\Ampol\\Data\\Output_v1.csv', index=False)






