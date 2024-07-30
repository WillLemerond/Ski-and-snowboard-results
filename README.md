import pandas as pd
import openpyxl
from openpyxl.styles import Border, Side, PatternFill, Alignment, Font


# Load the original CSV file
file_path = '/Users/williamlemerond/Desktop/untitled folder/PythonFile3.csv'
data = pd.read_csv(file_path)

# Rename the column "Comp.SetDetail" to "Category"
data = data.rename(columns={"Comp.SetDetail": "Category"})

# Delete all columns except for Date, Category, Sport, Discipline, Gender, Rank, Person, Age, Host City, and Host Country
columns_to_keep = ["Date", "Category", "Sport", "Discipline", "Gender", "Rank", "Person", "Age", "Host City", "Host Country"]
data = data[columns_to_keep]

# Delete any rows with a Category of "World Junior Championships" and a Rank above 10
data = data[~((data["Category"] == "World Junior Championships") & (data["Rank"] > 10))]

# Sort the data in the order of Date, Category, Sport, Discipline, Gender, Rank (least to most)
data = data.sort_values(by=["Date", "Category", "Sport", "Discipline", "Gender", "Rank"])

# Save the sorted data to an Excel file for formatting
excel_path = '/Users/williamlemerond/Desktop/untitled folder/PythonFile3.xlsx'
data.to_excel(excel_path, index=False)

# Load the Excel file for formatting
wb = openpyxl.load_workbook(excel_path)
ws = wb.active

# Define border styles
thick_border = Border(left=Side(style='thick'), 
                      right=Side(style='thick'), 
                      top=Side(style='thick'), 
                      bottom=Side(style='thick'))
thin_border = Border(bottom=Side(style='thin'))

# Change all applicable cells to a fill background of white and remove all borders
white_fill = PatternFill(start_color='FFFFFFFF', end_color='FFFFFFFF', fill_type='solid')
for row in ws.iter_rows(min_row=1, max_col=len(columns_to_keep)):
    for cell in row:
        cell.fill = white_fill
        cell.border = Border()

# Define a font style
header_font = Font(name='Arial', size=8, bold=True)
data_font = Font(name='Arial', size=8, bold=False)

# Apply font style to the header row and calculate the max width for each column
for cell in ws[1]:
    cell.font = header_font
    max_length = len(str(cell.value))
    column_letter = cell.column_letter
    
    # Find the maximum length in the column
    for data_cell in ws[column_letter]:
        try:
            max_length = max(max_length, len(str(data_cell.value)))
        except:
            pass
        data_cell.font = data_font if data_cell.row > 1 else header_font
        
    # Set the column width to the maximum length found, with a small margin
    adjusted_width = max_length + 2
    ws.column_dimensions[column_letter].width = adjusted_width
    

# Automatically adjust columns to fit the data
for column in ws.columns:
    max_length = 0
    column_letter = column[0].column_letter
    for cell in column:
        try:
            cell_length = len(str(cell.value))
            if cell_length > max_length:
                max_length = cell_length               
        except:
            pass
    adjusted_width = (max_length)
    ws.column_dimensions[column_letter].width = adjusted_width

# Move column names to the left, except Age and Rank
for cell in ws[1]:
    if cell.column_letter in ['H', 'F']:  # Age and Rank
        cell.alignment = Alignment(horizontal='center')
    else:
        cell.alignment = Alignment(horizontal='left')

# Align data in Age column to the center
for col in ws.iter_cols(min_row=2, max_row=ws.max_row, min_col=6, max_col=6):
    for cell in col:
        cell.alignment = Alignment(horizontal='center')

# Align data in Rank column to the center
for col in ws.iter_cols(min_row=2, max_row=ws.max_row, min_col=8, max_col=8):
    for cell in col:
        cell.alignment = Alignment(horizontal='center')
    
# Add thin borders to separate rows with different Sports, Disciplines, or Genders
previous_row = None
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
    if previous_row:
        if (row[2].value != previous_row[2].value or  # Sport
            row[3].value != previous_row[3].value or  # Discipline
            row[4].value != previous_row[4].value):  # Gender
            for cell in previous_row:
                cell.border = thin_border
    previous_row = row

# Add thick border to separate rows with different Dates
current_date = None
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
    if row[0].value != current_date:
        current_date = row[0].value
        for cell in row:
            cell.border = cell.border.copy(top=Side(style='thick'))

# Create a thick border on the bottom side of each cell with the column names
for cell in ws[1]:
    cell.border = cell.border.copy(bottom=Side(style='thick'))

# Create a thick border along the right side of each cell in the Host Country column
host_country_col_index = columns_to_keep.index("Host Country") + 1
for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=host_country_col_index, max_col=host_country_col_index):
    for cell in row:
        cell.border = cell.border.copy(right=Side(style='thick'))

# Create a thick border on the bottom side of each cell in the last row
for cell in ws[ws.max_row]:
    cell.border = cell.border.copy(bottom=Side(style='thick'))

# Create a thick border on the top side of each cell with the column name
for cell in ws[1]:
    cell.border = cell.border.copy(top=Side(style='thick'))

# Create a thick border along the left side of each cell in the Date column
for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=1):
    for cell in row:
        cell.border = cell.border.copy(left=Side(style='thick'))

# Highlight rows with a rank of 4-10 in powder blue, and highlight rows with a rank of 1-3 in pistachio
powder_blue_fill = PatternFill(start_color='B6D0E2', end_color='B6D0E2', fill_type='solid')
pistachio_fill = PatternFill(start_color='93C572', end_color='93C572', fill_type='solid')

for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
    rank = row[5].value  # Rank
    if 4 <= rank <= 10:
        for cell in row:
            cell.fill = powder_blue_fill
    elif 1 <= rank <= 3:
        for cell in row:
            cell.fill = pistachio_fill

# Save the formatted Excel file
wb.save(excel_path)
