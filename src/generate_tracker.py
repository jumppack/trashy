import pandas as pd
import os

def generate_tracker_final():
    # ---------------------------------------------------------
    # 1. SETUP THE DATA
    # ---------------------------------------------------------
    rows = 30
    # "Dent Method" brackets
    dent_bracket = "[     ]" 
    
    data = {
        "Turn #": list(range(1, rows + 1)),
        "KM": [dent_bracket] * rows,
        "NP": [dent_bracket] * rows,
        "PS": [dent_bracket] * rows,
        "LS": [dent_bracket] * rows
    }

    df = pd.DataFrame(data)

    # Add a "TOTALS" row at the bottom
    df.loc[rows] = ["TOTALS", "____", "____", "____", "____"]

    # ---------------------------------------------------------
    # 2. CREATE THE EXCEL FILE
    # ---------------------------------------------------------
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    filename = os.path.join(output_dir, 'Trash_Rotation_Printable.xlsx')
    
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    
    # Write data starting at row 3 (leaving space for Title/Instr)
    df.to_excel(writer, sheet_name='Sheet1', index=False, startrow=3)

    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # Hide default gridlines for a clean look
    worksheet.hide_gridlines(2)

    # ---------------------------------------------------------
    # 3. DEFINE STYLES
    # ---------------------------------------------------------
    # Main Body: Borders ONLY for the table
    cell_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1, # Thin border
        'font_name': 'Courier New',
        'font_size': 11
    })

    # Header Style (KM, NP...)
    header_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#D7E4BC',
        'border': 2, # Thick border
        'font_size': 14,
        'font_name': 'Courier New'
    })

    # Title Style (Full width)
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 18,
        'align': 'center',
        'valign': 'vcenter',
        'font_name': 'Arial'
    })

    # Instruction Style
    instr_format = workbook.add_format({
        'italic': True,
        'font_size': 10,
        'align': 'center',
        'font_name': 'Arial'
    })
    
    # Totals Row Style
    total_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'top': 2, # Thick top border
        'font_name': 'Courier New',
        'font_size': 12
    })

    # ---------------------------------------------------------
    # 4. FORMATTING & LAYOUT
    # ---------------------------------------------------------
    # Column Widths
    worksheet.set_column('A:A', 8, cell_format)
    worksheet.set_column('B:E', 20, cell_format)

    # Apply Header Format
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(3, col_num, value, header_format)

    # Apply Totals Row Format
    last_row = 3 + rows + 1
    worksheet.write(last_row, 0, "TOTALS", total_format)
    for col in range(1, 5):
        worksheet.write(last_row, col, "____", total_format)

    # Title & Instructions - Centered across A-E
    worksheet.merge_range('A1:E1', 'TRASH ROTATION TRACKER', title_format)
    worksheet.merge_range('A2:E2', 'FLOW: KM -> NP -> PS -> LS -> (Loop)', instr_format)

    # Footer Instructions
    footer_row = last_row + 2
    worksheet.merge_range(f'A{footer_row}:E{footer_row}', 
                          "INSTRUCTION: Find first empty box. Press Key firmly into box [     ].", 
                          instr_format)

    # ---------------------------------------------------------
    # 5. PRINT SETTINGS (CRITICAL FOR A4)
    # ---------------------------------------------------------
    worksheet.set_paper(9)  # A4
    worksheet.fit_to_pages(1, 1) # Fit all on one page
    worksheet.center_horizontally()
    worksheet.set_margins(left=0.5, right=0.5, top=0.5, bottom=0.5)

    writer.close()
    print(f"Success! File '{filename}' created.")

if __name__ == "__main__":
    generate_tracker_final()
