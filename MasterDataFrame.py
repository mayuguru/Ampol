import pandas as pd
print("---------- Reading Data file ------------")

df=pd.read_excel('C:\\GIT\\Ampol\\Data\\input1.xlsx',sheet_name='Sheet1',engine="openpyxl",header=None)

#print(df)

rowlength=len(df)
columnslength=len(df.columns)

dfhead=df.iloc[0].to_frame(name='col_0')

column_names = ["product_name", "product_hier_1", "product_hier_2", "product_hier_3", "year", "period_type", "period",
                "unit", "value"]

finaldf = pd.DataFrame(columns=column_names)

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