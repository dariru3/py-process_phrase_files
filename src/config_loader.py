CONFIG = {
    "GeneralSettings": { # When updating Colab, replace with commented folder paths
        "InputFolderPath": "data/input_files/", # "/content/drive/MyDrive/MagicBox/",
        "OutputFolderPath": "data/output_files/", # "/content/drive/MyDrive/MagicBox/Output_Folder/",
        "Column_Headers": ["ID", "Source", "Target", "Match", "Comment"]
    },
    "ProcessingDocSettings": {
        "DeleteFirstNTables": 3,
        "ColumnsToKeep": [0, 3, 5, 6, 7]
    },
    "ProcessingXliffSettings": {
        "TagPatterns": r"\{.?>|<.?\}|\{j\}", # Remove custom tags such as {b>, <b}, {j} from the input text.
        "XliffNamespace": "urn:oasis:names:tc:xliff:document:1.2",
    },
    "ConditionalFormattingSettings": {
        "TargetColumnIndex": 2,
        "MatchColumnIndex": 3,
        "CommentColumnIndex": 4,
        "CommentToGray": ["lock", "locked"],
        "MatchToGray": ["100", "101"],
        "BackgroundColor": "D9D9D9"
    },
    "TableFormattingSettings": {
        "RowWidths": [9, 81, 112, 11, 21],
        "NewColumnNames": {'ID': 'p', 'Source': 'Japanese', 'Target': 'English'}
    }
}
