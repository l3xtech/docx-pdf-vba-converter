# DOCX to PDF Converter with VBA

This repository contains a VBA (Visual Basic for Applications) script that converts Microsoft Word (.docx) files to Adobe PDF (.pdf) format using Microsoft Word. This script has been specifically created to be used in Windows environments that have severe restrictions on the installation of new programs and libraries.

## Requirements

This script requires an installation of Microsoft Word, as it uses Word's functionality to convert the documents. The script was developed and tested with Word 2016, but it should work with newer versions as well.

## Usage

1. Open Microsoft Word.
2. Press `Alt + F11` to open the VBA Editor.
3. In the VBA Editor, click on `Insert` > `Module` to create a new module.
4. Paste the content of the `ConvertDocxToPDF.bas` file into the module.
5. Modify the value of the `folderPath` variable to the directory that contains the .docx files you wish to convert.
6. Run the script by pressing `F5` or selecting `Run` > `Run Sub/UserForm`.

The script will convert all .docx files in the specified directory to .pdf files. Please remember that this will only work if Microsoft Word is installed on the machine where the script is being run.

## Warning

This script does not save changes made to .docx documents during the conversion. If you have documents that contain unsaved changes, please save these changes before running this script.

## Support

This script is provided as-is, with no warranties or promises of support. However, feel free to open an issue if you encounter problems or have questions.

