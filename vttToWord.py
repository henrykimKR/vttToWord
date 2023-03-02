# Title: Convert VTT to Word for Transcript
# Author: Henry Kim <Henry.Kim@gmail.com>
# Github: https://github.com/henrykimKR/vttToWord.git
# Date: 2023-03-01

import docx
import re
import os

# 1. List all the files in the current directory.
files = os.listdir()

# 2. Filter the list to only show .vtt files.
vtt_files = [file for file in files if file.endswith('.vtt')]

# 3. Display the list of .vtt files to the user.
print("Please select a .vtt file to convert.")
for i, file in enumerate(vtt_files):
    print(f"{i + 1}. {file}")

# 4. Ask the user to select a file and get the filename.
while True:
    try:
        choice = int(input("Enter the number of the file you want to convert: "))
        filename = vtt_files[choice - 1]
        with open(filename, 'r') as file:
            lines = file.readlines()
        break
    except (ValueError, IndexError):
        print("Invalid choice. Please try again.")

# 5. Create a new Word document with the same name as the vtt file.
doc_filename = os.path.splitext(filename)[0] + ".docx"
doc = docx.Document()

# 6. Add a table with two columns to the document.
table = doc.add_table(rows=0, cols=2)
table.style = 'Table Grid'

# 7. Remove the first line and any timestamps.
lines.pop(0)
lines = [line for line in lines if not re.match(r'^\d{2}:\d{2}:\d{2}.\d{3}.*', line)]

# 8. Remove the blank lines.
lines = [line for line in lines if line.strip()]

# 9. If there is a speaker's name within double quotation marks, put only the name (without the quotation marks) into the left column and remove the line.
for line in lines:
    if '"' in line:
        speaker = re.findall(r'"([^"]*)"', line)[0]
        table.add_row().cells[0].text = speaker
    # 10. Else if the line contains only numbers, remove the numbers and leave the left column blank.
    elif line.strip().isdigit():
        table.add_row().cells[0].text = ""
    # 11. Else, add it to the right column of the table.
    else:
        table.rows[-1].cells[1].text += line.strip()

doc.save(doc_filename)
