import pandas as pd
from process_word import process_word_file
from process_mxliff import parse_mxliff_to_df

word_input_file = "input_files/240226_良品計画様_統合報告2023対訳表_P37-40+-ja-en-T.docx"
word_output_folder = "output_files/"
mxliff_file = "input_files/240226_良品計画様_統合報告2023対訳表_P37-40+-ja-en-T.mxliff"

# Word file has comments
df1 = process_word_file(word_input_file, word_output_folder)
# .mxliff file has repeating text
df2 = parse_mxliff_to_df(mxliff_file)


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

    # Assuming the data structures you provided and focusing on completing the dataset as requested
    csv_file = 'output_files/merged_dfs.csv'
    df_combined.to_csv(csv_file, index=False)

    return df_combined


if __name__ == "__main__":
    merge_dfs(df1, df2)