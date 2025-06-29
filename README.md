# Compile_file_folderwise-

## Overview

This repository contains a VBA macro for Microsoft Excel that allows users to export the file structure of a selected folder into an Excel worksheet. The macro lists files in a folder (and its subfolders), presenting the hierarchy in columns, along with useful metadata for each file.

## Features

- **Folder Picker**: Lets users select a root folder for scanning.
- **Recursive Listing**: Traverses all subfolders and files within the selected directory.
- **Hierarchical Output**: Displays each folder level in its own Excel column.
- **File Metadata**: For each file, the output includes:
  - File name
  - File path
  - File size (KB)
  - Date modified
  - Direct hyperlink to open the file
- **Summary**: At the end, provides the total number of files and the maximum depth of the folder structure.
- **Clean Output**: Previous outputs in the worksheet are safely deleted before running a new scan.

## Usage

1. Open the VBA-enabled Excel workbook.
2. Press `Alt + F11` to open the VBA Editor.
3. Import the `1.FolderStructureToColumns_Improved_noFolderHyperlink_Version2.bas` file if it isn't already present.
4. Run the `FolderStructureToColumns_Improved` macro.
5. When prompted, select the root folder you wish to scan.
6. The macro will generate a new worksheet named "Folder Structure Columns" containing the results.

## Example Output

| Level 1 | Level 2 | ... | File Name | File Path | File Size (KB) | Date Modified | Hyperlink   |
|---------|---------|-----|-----------|-----------|----------------|---------------|-------------|
| FolderA | SubA    |     | file1.txt | ...       | 10.5           | ...           | Open File   |
| ...     | ...     | ... | ...       | ...       | ...            | ...           | ...         |

## Requirements

- Microsoft Excel with macro support (VBA).
- Windows OS (requires FileSystemObject and FileDialog).

## License

_No license specified. Please add one if required._

---

Feel free to modify the macro or README as needed to fit your workflow!
