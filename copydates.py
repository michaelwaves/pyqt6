import pandas as pd
from math import isnan

FILENAME = 'data.xlsx'

df = pd.read_excel(FILENAME)
dates = df['Date'].tolist()
fps = df['FP'].tolist()
fys = df['FY'].tolist()
voucher_categories= df['Voucher Category'].tolist()
voucher_no = df['Voucher No.'].tolist()

def fill_blanks(list):
    current_element = list
    for index, value in enumerate(list):
        #if it is a datetime, isnan will fail and except branch is triggered.
        try:
            if isnan(value):
                list[index] = current_element
            else:
                current_element = value  
        except:
            current_element = value
    return list

new_dates = fill_blanks(dates)
new_fps = fill_blanks(fps)
new_fys = fill_blanks(fys)
new_voucher_categories = fill_blanks(voucher_categories)
new_voucher_no = fill_blanks(voucher_no)

df['Date'] = dates
df['FP'] = fps
df['FY'] = fys
df['Voucher Category'] = voucher_categories
df['Voucher No.'] = voucher_no

df.to_excel('output.xlsx')