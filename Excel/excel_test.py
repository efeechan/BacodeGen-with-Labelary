import pandas as pd
import re
from datetime import datetime

# Read the Excel file
excel_data = pd.read_excel(
    "C:/Users/eyilmazdemir/Desktop/barcode/excel/nestle.xlsx", dtype=str
)

# Get the number of rows and columns in the Excel file
num_rows, num_columns = excel_data.shape

# Create an empty list for each line
lines = []

# Iterate over each row in the Excel file
for i in range(num_rows):
    line = []
    for j in range(num_columns):
        if pd.notna(excel_data.iloc[i, j]):
            line.append(excel_data.iloc[i, j])
    lines.append(line)


# Define the function to convert the date format
def convert_date_format(date_str):
    return datetime.strptime(date_str, "%d.%m.%Y").strftime("%y%m%d")


# Get the converted dates from the Excel file
converted_dates = [
    convert_date_format(date_str) for date_str in excel_data.iloc[:, 0].dropna()
]


# Define the function to replace a value in a line
def replaceValzpl(OldLine, start_char, end_char, new_value):
    pattern = re.compile(re.escape(start_char) + r"(.*?)" + re.escape(end_char))
    replaced_line = re.sub(pattern, start_char + new_value + end_char, OldLine)
    return replaced_line


# Read the ZPL file
with open("C:/Users/eyilmazdemir/Desktop/barcode/input/sartenestle.zpl", "r") as file:
    zpl_content = file.readlines()

# Debugging print statements
# print("Total lines in Excel file:", num_rows)
# print("First row in Excel data:", lines[0])
# print("Converted dates:", converted_dates)

# Iterate over each line in the Excel data and modify the ZPL content accordingly
for i in range(len(lines)):
    # print(f"\nProcessing line {i+1} from the Excel file with data:", lines)

    # Modify the desired lines using the data from the current Excel line
    modified_zpl_content = (
        zpl_content.copy()
    )  # Make a copy of the original ZPL content to modify
    modified_zpl_content[33] = replaceValzpl(
        modified_zpl_content[33],
        "^FD>;>8",
        "^FS",
        lines[i][20]
        + lines[i][21]
        + lines[i][22]
        + converted_dates[i]
        + lines[i][26]
        + lines[i][13]
        + ">8"
        + lines[i][24]
        + lines[i][17],
    )
    modified_zpl_content[34] = (
        replaceValzpl(
            modified_zpl_content[34],
            "^FD",
            "^FS",
            "("
            + lines[i][20]
            + ") "
            + (lines[i][21])[0:3]
            + " "
            + (lines[i][21])[3:8]
            + " "
            + (lines[i][21])[8:13]
            + " "
            + (lines[i][21])[13:14]
            + " ("
            + lines[i][22]
            + ") "
            + converted_dates[i]
            + " ("
            + lines[i][26]
            + ") "
            + lines[i][13]
            + " ("
            + lines[i][24]
            + ") "
            + lines[i][17],
        ),
    )
    modified_zpl_content[10] = replaceValzpl(
        modified_zpl_content[10], "^FD", "^FS", lines[i][1]
    )
    modified_zpl_content[9] = replaceValzpl(
        modified_zpl_content[9], "^FD", "^FS", lines[i][0]
    )
    modified_zpl_content[7] = replaceValzpl(
        modified_zpl_content[7], "^FD", "^FS", lines[i][2]
    )
    modified_zpl_content[6] = replaceValzpl(
        modified_zpl_content[6], "3-", "^FS", lines[i][3]
    )
    modified_zpl_content[16] = replaceValzpl(
        modified_zpl_content[16], "3-", "^FS", lines[i][3]
    )
    modified_zpl_content[21] = replaceValzpl(
        modified_zpl_content[21], "^FD", "^FS", lines[i][3]
    )
    modified_zpl_content[5] = replaceValzpl(
        modified_zpl_content[5], "^FD", "^FS", lines[i][4]
    )
    modified_zpl_content[8] = replaceValzpl(
        modified_zpl_content[8], "^FD", "^FS", lines[i][5]
    )
    modified_zpl_content[57] = replaceValzpl(
        modified_zpl_content[57], "^FD", "^FS", lines[i][7]
    )
    modified_zpl_content[24] = replaceValzpl(
        modified_zpl_content[24], "^FD", "^FS", lines[i][9]
    )
    modified_zpl_content[25] = replaceValzpl(
        modified_zpl_content[25], "^FD", "^FS", lines[i][10]
    )
    modified_zpl_content[31] = replaceValzpl(
        modified_zpl_content[31], "^FD", " ", lines[i][11]
    )
    modified_zpl_content[31] = replaceValzpl(
        modified_zpl_content[31], " ", "^FS", lines[i][12]
    )
    modified_zpl_content[49] = replaceValzpl(
        modified_zpl_content[49], "^FD", "^FS", lines[i][14]
    )
    modified_zpl_content[51] = replaceValzpl(
        modified_zpl_content[51], "^FD", "^FS", lines[i][15]
    )
    modified_zpl_content[55] = replaceValzpl(
        modified_zpl_content[55], "^FD", "^FS", lines[i][16]
    )
    modified_zpl_content[36] = replaceValzpl(
        modified_zpl_content[36], "^FD>;>8", "^FS", lines[i][18]
    )
    modified_zpl_content[40] = replaceValzpl(
        modified_zpl_content[40], "^FD", "^FS", lines[i][18]
    )
    modified_zpl_content[45] = replaceValzpl(
        modified_zpl_content[45], "^FD>;>200", "^FS", lines[i][13]
    )
    modified_zpl_content[46] = replaceValzpl(
        modified_zpl_content[46], "^FD200", "^FS", lines[i][13]
    )
    modified_zpl_content[53] = replaceValzpl(
        modified_zpl_content[53], "^FD200", "^FS", lines[i][13]
    )
    modified_zpl_content[13] = replaceValzpl(
        modified_zpl_content[13],
        "^FD",
        "^FS",
        lines[i][17][1:3] + "." + lines[i][17][3:6],
    )
    modified_zpl_content[37] = replaceValzpl(
        modified_zpl_content[37],
        "^FD",
        "^FS",
        "("
        + lines[i][18][0:2]
        + ") "
        + lines[i][18][2:3]
        + " "
        +lines[i][18][3:5]
        + " "
        +lines[i][18][5:10]
        + " "
        +lines[i][18][10:13]
        + " "
        +lines[i][18][13:16]
        + " "
        +lines[i][18][16:19]
        + " "
        +lines[i][18][19:20],
    )
    modified_zpl_content[17] = replaceValzpl(
        modified_zpl_content[17], "3-", "^FS", lines[i][3]
    )
    # Add more modifications based on the current line from Excel, as needed

    # Assuming modified_zpl_content is a list of tuples, convert it to a list of strings
    modified_zpl_content = ["".join(item) for item in modified_zpl_content]

    # Save the modified ZPL content to a new file for the current Excel line
    output_file_path = f"C:/Users/eyilmazdemir/Desktop/barcode/input/{lines[i][8]}.zpl"
    with open(output_file_path, "w") as output_file:
        output_file.writelines(modified_zpl_content)

    # print(f"ZPL file for line {i+1} saved as: {output_file_path}")

# print("\nZPL files modified and saved for each line in the Excel data.")
