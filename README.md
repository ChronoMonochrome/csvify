# Python Office Utils
Python utils to work with Office documents

# docx2csv.py
Usage:
docx2csv.py input_dir [output_dir]

For each found docx file in input_dir directory,
script will try to recursively find tables containing a 
[caption](https://support.office.com/en-us/article/add-format-or-delete-captions-in-word-82fa82a4-f0f3-438f-a422-34bb5cef9c81).
For each such found table a CSV file will be produced in the output_folder
(default is <input_dir>/out if output_folder is not specified), keeping the original folder structure.
