CONFIG = {
    "GeneralSettings": { # When updating Colab, replace with commented folder paths
        "InputFolderPath": "data/input_files/", # "/content/drive/MyDrive/MagicBox/",
        "OutputFolderPath": "data/output_files/", # "/content/drive/MyDrive/MagicBox/Output_Folder/",
        "Column_Headers": ["ID", "Source", "Target", "Match", "Comment"]
    },
    "ProcessingSettings": {
        "DeleteFirstNTables": 3,
        "MaxAttempts": 2,
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
        "RowWidths": [9, 81, 112, 11, 21]
    }
}
