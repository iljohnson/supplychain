# python program to run SliverMeal algorithm which allows users to input the starting inventory level and adjust set-up
# costs and holding costs for each period by importing the starting data in a data frame format (excel spreadsheet .


import pandas as pd
import os.path

desired_width = 320
pd.set_option('display.width', desired_width)

pd.set_option('display.max_columns', 20)

homepath = os.path.expanduser("~\\Desktop")

open_file = os.path.join(homepath, 'SilverMeal_Test_Data.xlsx')

# print(open_file)

df = pd.read_excel(open_file, sheet_name='Sheet1')

Int_Inv_Hold = input("What is your initial inventory holding   ") or 0

df['Inv_On_Hand'] = 0
df['Inv_Carry_fwd'] = 0
df['Make'] = 0
df['Make_Qty'] = 0
df['Accum_Make'] = 0
df['Man_Cost'] = 0
df['Hld_Cost'] = 0
df['Z_Costs'] = 0
df['SM_Calc'] = 0


P = 0
N = 0
M = 0
RT = 0              # running total of Z_Costs per manufacturing period
RT2 = 0             # running total of accum_make qty per manufacturing period


while P + N <= df['Period'].max()-1:

    if df.iloc[N+P, 0] == 1:
        df.iloc[N+P, 4] = int(Int_Inv_Hold)                         # Inv_On_Hand
    else:
        df.iloc[N+P, 4] = df.iloc[N+P-1, 5]

    if df.iloc[N+P, 4]-df.iloc[N+P, 1] > 0:
        df.iloc[N+P, 5] = df.iloc[N+P, 4]-df.iloc[N+P, 1]
    else:
        df.iloc[N+P, 5] = 0                                         # Inventory Carry Forward

    if df.iloc[N+P, 1] > df.iloc[N+P, 4]:
        df.iloc[N+P, 7] = df.iloc[N+P, 1]-df.iloc[N+P, 4]
    else:
        0                                                          # Make_Qty

    RT2 = RT2 + df.iloc[N+P, 7]                                    # Accum_make
    df.iloc[N+P, 8] = RT2

    if df.iloc[N+P, 8] == df.iloc[N + P, 7] > 0:
        df.iloc[N+P, 6] = 1
        M = 1
    else:
        0
        M = 0                                                       # Make

    df.iloc[N+P, 9] = df.iloc[N+P, 3] * M                           # Man_Cost
    df.iloc[N+P, 10] = df.iloc[N+P, 2] * df.iloc[N+P, 7] * N + df.iloc[N+P, 2] * df.iloc[N+P, 5]  # Hold Costs
    # print(N)
    df.iloc[N+P, 11] = df.iloc[N+P, 9] + df.iloc[N+P, 10]           # Z_Costs Sum of Costs

    RT = RT + df.iloc[N+P, 11]

    df.iloc[N+P, 12] = round(RT/(N+1), 2)                           # Average Costs per period for SM calculation

    # df['N_Track'] = N

    # print(df)

    # loop after first run and compare after the second run

    if df.iloc[N+P, 12] > df.iloc[N+P-1, 12] and N != 0 and df.iloc[N+P, 7] > 0:
        P = N + P
        N = 0
        RT = 0
        RT2 = 0

    else:
        N += 1


total = sum(df.Z_Costs)
Tot_Man_Costs = sum(df.Man_Cost)
Tot_Hold_Costs = sum(df.Hld_Cost)

df = df.append({'Man_Cost': Tot_Man_Costs, 'Hld_Cost': Tot_Hold_Costs, 'Z_Costs': total}, ignore_index=True)

print(df)
#  print("Total ", total, "Total Manufacturing Costs ", Tot_Man_Costs,"Total Holding Costs ", Tot_Hold_Costs)

export_file = os.path.join(homepath, 'SilverMeal_results.xlsx')

export_to_excel = df.to_excel(export_file, index=False, header=True)

