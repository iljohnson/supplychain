import pandas as pd
import os.path

desired_width = 320
pd.set_option('display.width', desired_width)

pd.set_option('display.max_columns', 20)

homepath = os.path.expanduser("~\\Desktop")

open_file = os.path.join(homepath, 'WagnerWhiten_Test_Data.xlsx')

# print(open_file)

df = pd.read_excel(open_file, sheet_name='Sheet1')

Hold_Cost = float(input("What is your inventory holding cost?  ")) or float(1)

Setup_Cost = float(input("What is your set-up cost?  "))or float(1)

df_shape = df.shape

Max_Row = df_shape[0]

Max_Col = df_shape[1]


df.iloc[0, 1] = Setup_Cost  # row then col

row_index = 0
col_index = 2
col_index_start = 1
Min_cost = 0
RT = 0

while row_index < Max_Row:

    while col_index < Max_Col:
        if row_index == 0:
            Min_cost = Setup_Cost
            RT = RT + df.iloc[Max_Row-2, col_index] * (df.iloc[Max_Row - 1, col_index] - col_index_start)
            df.iloc[row_index, col_index] = RT + Min_cost
            col_index += 1
            # one time buy
        else:
            Min_cost = df.iloc[0:Max_Row - 2, col_index_start-1].min()
            RT = RT + df.iloc[Max_Row - 2, col_index] * (df.iloc[Max_Row - 1, col_index] - col_index_start)
            df.iloc[row_index, col_index] = RT + Min_cost
            col_index += 1
    row_index += 1
    col_index_start += 1
    col_index = col_index_start
    Min_cost = 0
    RT = Setup_Cost


print(df)


min_idx = Max_Col-1

task = ""

while min_idx > 0:
    min_idx_old = min_idx
    min_value = df.iloc[0:Max_Row - 2, min_idx_old].min()
    min_idx = df.index[df.iloc[0:Max_Row - 2, min_idx_old].idxmin(axis=1)]
    make_qty = str(df.iloc[12, min_idx+1:min_idx_old + 1].sum())
    task = ("In period " + str(min_idx + 1) + " make quantity " + make_qty + "   ") + task

print(task)

open_my_file = os.path.join(homepath, 'WagnerWhiten_Results.xlsx')

writer = pd.ExcelWriter(open_my_file)
df.to_excel(writer, startrow=4, startcol=0)

worksheet = writer.sheets['Sheet1']
worksheet.write(0, 0, task)
worksheet.write(1, 0, "Set-up cost = " + str(Setup_Cost))
worksheet.write(2, 0, "Holding cost = " + str(Hold_Cost))

writer.save()
