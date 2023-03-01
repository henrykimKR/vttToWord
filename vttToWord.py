# Title: Convert VTT to Word for Transcript
# Author: Henry Kim <Henry.Kim@gmail.com>
# Date: 2023-02-28

import docx
import re

# 1. Extract lines from a .vtt file.
filename = input("Enter the filename of the VTT file: ")
with open(filename, 'r') as file:
    lines = file.readlines()

# 2. Create a new Word document and add a table with two columns.
doc = docx.Document()
table = doc.add_table(rows=0, cols=2)
table.style = 'Table Grid'

# 3. Remove the first line and any timestamps.
lines.pop(0)
lines = [line for line in lines if not re.match(r'^\d{2}:\d{2}:\d{2}.\d{3}.*', line)]

# 4. Remove the blank lines.
lines = [line for line in lines if line.strip()]

# 5. If there is a speaker's name within double quotation marks, put only the name (without the quotation marks) into the left column and remove the line.
for line in lines:
    if '"' in line:
        speaker = re.findall(r'"([^"]*)"', line)[0]
        table.add_row().cells[0].text = speaker
    # 6. Else if the line contains only numbers, remove the numbers and leave the left column blank.
    elif line.strip().isdigit():
        table.add_row().cells[0].text = ""
    # 7. Else, add it to the right column of the table.
    else:
        table.rows[-1].cells[1].text += line.strip()

doc.save('output.docx')
