import pandas as pd

def merge_dfs(df1, df2):
    print("========== Initial df1 ==========")
    print(df1)
    print("\n========== Initial df2 ==========")
    print(df2)

    # Convert Index to int
    df1['Index'] = df1['Index'].astype(int)
    df2['Index'] = df2['Index'].astype(int)
    print("\n========== After converting Index to int ==========")
    print("df1:")
    print(df1)
    print("df2:")
    print(df2)

    # Convert Match to int, with error handling
    df1['Match'] = pd.to_numeric(df1['Match'], errors='coerce').fillna(0).astype(int)
    df2['Match'] = pd.to_numeric(df2['Match'], errors='coerce').fillna(0).astype(int)
    print("\n========== After converting Match to int ==========")
    print("df1:")
    print(df1)
    print("df2:")
    print(df2)

    # Merge the DataFrames
    df_combined = pd.merge(df1, df2, on=['Index', 'Source'], how='outer', suffixes=('', '_df2'))
    print("\n========== After merging df1 and df2 ==========")
    print(df_combined)

    # Select the best values for each column based on availability and preference
    df_combined['Target'] = df_combined['Target'].where(df_combined['Target'] != '', df_combined['Target_df2'])
    print("\n========== After selecting best Target values ==========")
    print(df_combined[['Index', 'Source', 'Target', 'Target_df2']])

    df_combined['Match'] = df_combined['Match'].fillna(0).astype(int)
    df_combined['Match'] = df_combined['Match'].where(df_combined['Match'] != 0, df_combined['Match_df2']).fillna(0).astype(int)
    print("\n========== After selecting best Match values ==========")
    print(df_combined[['Index', 'Source', 'Match', 'Match_df2']])

    df_combined['Comment'] = df_combined['Comment'].where(df_combined['Comment'] != '', df_combined['Comment_df2'])
    print("\n========== After selecting best Comment values ==========")
    print(df_combined[['Index', 'Source', 'Comment', 'Comment_df2']])

    # Drop the temporary columns from df2
    df_combined.drop(columns=['Target_df2', 'Match_df2', 'Comment_df2'], inplace=True)
    print("\n========== Final combined DataFrame after dropping df2 temporary columns ==========")
    print(df_combined)

    return df_combined
