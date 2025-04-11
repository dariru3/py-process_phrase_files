import pandas as pd

def merge_dfs(df_word, df_mxliff):
    # Convert Match to int, with error handling
    df_word['Match'] = pd.to_numeric(df_word['Match'], errors='coerce').fillna(0).astype(int)
    df_mxliff['Match'] = pd.to_numeric(df_mxliff['Match'], errors='coerce').fillna(0).astype(int)

    # Merge the DataFrames
    df_combined = pd.merge(df_word, df_mxliff, on=['ID'], how='outer', suffixes=('', '_df2'))

    # Select the best values for each column based on availability and preference
    df_combined['Target'] = df_combined['Target'].where(df_combined['Target'] != '', df_combined['Target_df2'])

    df_combined['Match'] = df_combined['Match'].fillna(0).astype(int)
    df_combined['Match'] = df_combined['Match'].where(df_combined['Match'] != 0, df_combined['Match_df2']).fillna(0).astype(int)

    df_combined['Comment'] = df_combined['Comment'].where(df_combined['Comment'] != '', df_combined['Comment_df2'])

    # Drop the temporary columns from df2
    df_combined.drop(columns=['Source_df2', 'Target_df2', 'Match_df2', 'Comment_df2'], inplace=True)

    return df_combined
