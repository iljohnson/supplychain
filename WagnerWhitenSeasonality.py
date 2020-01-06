import pandas as pd
import os.path

desired_width = 320
pd.set_option('display.width', desired_width)

pd.set_option('display.max_columns', 20)

homepath = os.path.expanduser("~\\Desktop")

open_file = os.path.join(homepath, 'WagnerWhiten_Test_Data_Seasonality.xlsx')


df = pd.read_excel(open_file, sheet_name='Sheet2')

df_shape = df.shape

Max_Row = df_shape[0]

Max_Col = df_shape[1]


df.iloc[0, 1] = df.iloc[Max_Row-2, 1]

row_index = 0
col_index = 2
col_index_start = 1
Min_cost = 0
RT = 0

while row_index < Max_Row:

    while col_index < Max_Col:
        if row_index == 0:
            Min_cost = df.iloc[0, 1]
            RT = RT + df.iloc[Max_Row-4, col_index] * (df.iloc[Max_Row - 3, col_index] - col_index_start) * \
                df.iloc[Max_Row - 1, col_index]
            df.iloc[row_index, col_index] = RT + Min_cost
            col_index += 1
            # one time buy
        else:
            Min_cost = df.iloc[0:Max_Row - 4, col_index_start-1].min()
            RT = RT + df.iloc[Max_Row - 4, col_index] * (df.iloc[Max_Row - 3, col_index] - col_index_start) * \
                df.iloc[Max_Row - 1, col_index]
            df.iloc[row_index, col_index] = RT + Min_cost
            col_index += 1
    row_index += 1
    col_index_start += 1
    col_index = col_index_start
    Min_cost = 0
    RT = df.iloc[Max_Row-2, 2]


print(df)


min_idx = Max_Col-1

task = ""

while min_idx > 0:
    min_idx_old = min_idx
    min_value = df.iloc[0:Max_Row - 4, min_idx_old].min()
    min_idx = df.index[df.iloc[0:Max_Row - 4, min_idx_old].idxmin(axis=1)]
    make_qty = str(df.iloc[Max_Row-4, min_idx+1:min_idx_old + 1].sum())
    task = ("In period " + str(min_idx + 1) + " make quantity " + make_qty + "   ") + task

print(task)

open_my_file = os.path.join(homepath, 'WagnerWhiten_Seasonality_Results.xlsx')

writer = pd.ExcelWriter(open_my_file)
df.to_excel(writer, startrow=4, startcol=0)

worksheet = writer.sheets['Sheet1']
worksheet.write(0, 0, task)

writer.save()
