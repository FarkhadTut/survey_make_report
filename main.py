import pandas as pd
from glob import glob


files = {
        2022: 'out\\output_Сурхондарё_индикаторлар_номинал_20_11_2023_2022_01_01_2022_09_01.xlsx',
        2023: 'out\\output_Сурхондарё_индикаторлар_номинал_20_11_2023_2023_01_01_2023_09_01.xlsx'
        }

years = list(files.keys())

import pandas as pd

# Assuming you have two DataFrames: excel1 and excel2

# Read Excel files into DataFrames
excel1 = pd.read_excel(files[2022])
excel2 = pd.read_excel(files[2023])
columns_merged = []
for i, c in enumerate(excel1.columns.values.tolist()[5:]):
    columns_merged.append(c + "_x")
    columns_merged.append(excel2.columns.values.tolist()[i+5] + "_y")

columns_merged = excel1.columns.values.tolist()[:5] + columns_merged

merge_on = excel1.columns.values.tolist()[:5]
# Merge DataFrames on common columns
merged_df = pd.merge(excel1, excel2, on=merge_on)
print(merged_df.columns)
# Create a multi-level column structure
multi_level_columns = pd.MultiIndex.from_product([excel1.columns.values.tolist()[5:], [2022, 2023]], names=['Column', 'Year'])
print(multi_level_columns)
# Set the multi-level columns to the merged DataFrame
merged_df = merged_df[columns_merged]
merged_df.set_index(merge_on, inplace=True)

## dangerous
merged_df.columns = multi_level_columns


merged_df.reset_index(drop=False, inplace=True)
# Save the merged DataFrame to an Excel file with multi-level columns
merged_df.to_excel('path_to_excel3.xlsx', index=True)