# 📄 docx2md PowerShell Script

A PowerShell script that uses [Pandoc](https://pandoc.org) to convert `.docx` files to Markdown, removing empty lines for a cleaner Markdown format. This tool is especially useful for importing Word notes into Markdown-based programs like Obsidian or Notion.

---

## ✨ Features
- **Convert `.docx` to `.md`**: Generates clean, UTF-8 encoded Markdown files from Word documents.
- **Removes Empty Lines**: Cleans up unnecessary empty lines in the Markdown output.
- **Overwrite Protection**: Checks if the target file exists and prompts for overwrite confirmation.
- **Error Handling**: Catches common issues like incorrect file types, missing Pandoc installation, and invalid paths.

## 🔧 Prerequisites
- **[Pandoc](https://pandoc.org)**: Make sure Pandoc is installed and accessible in your system’s PATH.
- **PowerShell**: This script runs in PowerShell.

## 🚀 Setup and Usage

1. **Clone the repository**:
   ```sh
   git clone https://github.com/urban05/docx2md.git
   ```
   
2. **Add the `docx2md` alias to your PowerShell profile** for easy access. 
If trying it out temporarily, you can simply create a function:
   ```powershell
   function docx2md { .\path\to\docx2md.ps1 $args }
   ```
Keep in mind, that functions are lost when closing the shell

3. **Run the script**:
   - **With Arguments**:
     ```powershell
     docx2md -Source "C:\path\to\file.docx" -Target "C:\path\to\output.md"
     ```
   - **Without Arguments**: You’ll be prompted for the source and target paths interactively.

## ⚙️ Parameters
- **`-Source`**: Path to the `.docx` file you wish to convert.
- **`-Target`**: Path for saving the converted `.md` file.

If no parameters are specified, the script will ask for the source and target paths interactively.

## 🛠️ Error Handling
- If the `.docx` file is invalid or corrupted, the script will display an error and stop.
- If the target file already exists, the script prompts to confirm overwriting.

## 📚 Example
```powershell
docx2md -Source "C:\Documents\Notes.docx" -Target "C:\Documents\Notes.md"
```

## 📄 License
This project is licensed under the MIT License. See `LICENSE` for more information.
