![](https://i.imgur.com/KijwQNZ.png)

## Summary

Airtable is a relational database management system. This is a excel VBA script for downloading and renaming images off a CSV export.

Read more about it here - https://community.airtable.com/t/bulk-downloader-and-renamer-using-excel-vba/4913

### Prerequisites

You'll need to have microsoft excel installed on your computer. Windows OS preferred for running batch file commands.

## Installing macro from repo

You can run this macro by going to the "DataStager.VBS" file located in this repo.

1. Copy the contents of the file
2. Download and Open your excel file from Airtable if you haven’t (don’t copypaste cells, use the builtin -download CSV button)
3. Press ALT+F11
4. Insert code into a module
5. Run the submodule airtableCleaner()

## Installing macro from a excel .xlsm file

Alternatively, you can just download my .xlsm file here

https://www.dropbox.com/s/k1o16u28sj0e0l3/AirtableBulkDownloaderRenamer.xlsm?dl=0

1. just download my macro enabled workbook
2. Paste data in the first sheet (exactly the data you’d get from download as CSV from airtable)
3. Home → view → view macro → airtableCleaner() → run

## IMPORTANT NOTES BEFORE RUNNING MACRO!

All of Column A should be unique values. This macro doesn't run otherwise, due to potential naming conflicts

If you want to keep the original file name uploaded to amazonS3, you can easily keep a copy of your original column somewhere else.

You can use this macro for things besides airtable, just format it in the same original input as above.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details


