# Convert NB and DOC Files to PDF and Delete Originals

A PowerShell script that recursively searches a specified folder and all of its subfolders for supported document types, converts them to PDF, and automatically deletes the original files after successful conversion.

This script was originally developed to facilitate the archival processing of a large collection of Nota Bene documents by converting them into a more accessible and preservation-friendly PDF format.

---

## Features

- Converts the following file types to PDF:
  - `.nb` (Nota Bene)
  - `.docx`
  - `.doc`
  - `.docs`
- Searches all subfolders automatically.
- Ignores Microsoft Word temporary files (`~$`).
- Verifies successful PDF creation before deleting the original file.
- Keeps the original file if the conversion fails.
- Processes an entire directory with a single command.

---

## Requirements

- Windows
- Microsoft PowerShell
- Microsoft Word (installed)
- Read access to the source files
- Write permission in the destination folders

---

## How to Use

### Step 1

Open the PowerShell script and change the folder path to the directory you want to process.

```powershell
$folderPath = "C:\Users\Rajendra\Desktop\FolderName"
```

### Step 2

Save the script.

### Step 3

Run the script in PowerShell.

The script will automatically:

1. Search the selected folder and all subfolders.
2. Find all supported document types.
3. Convert each document to PDF.
4. Verify that the PDF was successfully created.
5. Delete the original file only after successful conversion.

---

## Supported File Types

| File Type | Supported |
|-----------|-----------|
| `.nb` | ✓ |
| `.docx` | ✓ |
| `.doc` | ✓ |
| `.docs` | ✓ |

---

## Important Notes

- Existing PDFs with the same filename may be overwritten.
- Temporary Microsoft Word files beginning with `~$` are ignored.
- If a conversion fails, the original file is preserved.
- The script searches every subfolder beneath the specified directory.

---

## Warning

⚠️ **This script permanently deletes the original files after successful PDF conversion.**

Before processing an important collection, it is strongly recommended that you:

- create a backup of your files, and
- test the script on a small sample folder first.

---

## License

This project is released under the MIT License.

---

## Author

**Rajendra Senchurey**

PhD Student  
School of Government and Public Policy  
The University of Arizona
