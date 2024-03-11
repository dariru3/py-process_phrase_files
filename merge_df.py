import pandas as pd

def merge_dfs(df1, df2):
    df1['Index'] = df1['Index'].astype(int)  # Convert Index to int
    df1['Match'] = pd.to_numeric(df1['Match'], errors='coerce').fillna(0).astype(int)  # Convert Match to int, with error handling

    df_combined = pd.merge(df1, df2, on=['Index', 'Source'], how='outer', suffixes=('', '_df2'))

    # Now, select the best values for each column based on availability and preference
    df_combined['Target'] = df_combined['Target'].where(df_combined['Target'] != '', df_combined['Target_df2'])
    df_combined['Match'] = df_combined['Match'].where(df_combined['Match'] != 0, df_combined['Match_df2'])
    df_combined['Comment'] = df_combined['Comment'].where(df_combined['Comment'] != '', df_combined['Comment_df2'])

    # Drop the temporary columns from df2
    df_combined.drop(columns=['Target_df2', 'Match_df2', 'Comment_df2'], inplace=True)

    return df_combined