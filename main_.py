import pandas as pd
from glob import glob


files = {
        2022: 'data\\Сурхондарё_индикаторлар_номинал_22_11_2023_2022_01_01_2022_09_01.xlsx',
        2023: 'data\\Сурхондарё_индикаторлар_номинал_22_11_2023_2023_01_01_2023_09_01.xlsx'
        }

def merged_table(files):
        dfs = {}
        for year, file in files.items():
                dfs[year] = pd.read_excel(file)

        years = sorted(list(files.keys()))
        all_columns = pd.read_excel(files[2022]).columns.values.tolist()
        columns = all_columns[5:]
        index_cols = all_columns[:5]

        multi_level_columns = pd.MultiIndex.from_product([columns, years + ["ўсиш, %"]], names=['Column', 'Year'])

        df_out = pd.DataFrame()
        for c in columns:
                for year, df in dfs.items():
                        df = df[index_cols + [c]]
                        df.rename(columns={c: c + f"_{year}"}, inplace=True)
                        if df_out.empty:
                                df_out = df
                        else:
                                df_out = pd.merge(df_out, df, on=index_cols)
                
                diff_col = c + "_diff%"
                new_year_col = c + f"_{years[1]}"
                old_year_col = c + f"_{years[0]}"
                df_out[diff_col] = df_out[new_year_col].div(df_out[old_year_col]).multiply(100)


        df_out.set_index(index_cols, inplace=True)
        df_out.columns = multi_level_columns
        df_out.to_excel('out\\merged\\merged_{years}.xlsx'.format(years="_".join([str(y) for y in years])))
        return df_out


merged_table(files)



                        
        



