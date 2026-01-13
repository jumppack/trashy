import pandas as pd
import os

def generate_tracker_v2():
    # ---------------------------------------------------------
    # 1. SETUP THE DATA
    # ---------------------------------------------------------
    # Goal: A high-density grid for the "Dent Method".
    # 30 Rows, 4 Columns (KM, NP, PS, LS).
    # Logic: Left -> Right, then Down.
    
    rows = 30
    columns = ["KM", "NP", "PS", "LS"]
    
    # The Dent Method requires thick brackets for key pressing
    dent_bracket = "[     ]" 
    
    data = {
        "Turn #": list(range(1, rows + 1)),
        "KM": [dent_bracket] * rows,
        "NP": [dent_bracket] * rows,
        "PS": [dent_bracket] * rows,
        "LS": [dent_bracket] * rows
    }

    df = pd.DataFrame(data)

    # Add a "TOTALS" row at the bottom for accountability
    df.loc[rows] = ["TOTALS", "____", "____", "____", "____"]

    # ---------------------------------------------------------
    # 2. CREATE THE EXCEL FILE
    # ---------------------------------------------------------
    output_dir = "out_v2"
    os.makedirs(output_dir, exist_ok=True)
    filename = os.path.join(output_dir, 'Trash_Rotation_Printable_v2.xlsx')
    
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    
    # Write data starting at row 3
    df.to_excel(writer, sheet_name='Sheet1', index=False, startrow=3)

    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # ---------------------------------------------------------
    # 3. DEFINE STYLES
    # ---------------------------------------------------------
    # Main Body: Monospace font for perfect alignment
    cell_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'font_name': 'Courier New',
        'font_size': 11
    })

    # Header Style
    header_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#D7E4BC',
        'border': 2, # Thick border
        'font_size': 14,
        'font_name': 'Courier New'
    })

    # Title Style
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
    worksheet.set_column('B:E', 20, cell_format) # Wider for dent brackets

    # Apply Header Format
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(3, col_num, value, header_format)

    # Apply Totals Row Format
    totals_row_idx = 3 + rows + 1 # Header row + data rows + 1 (0-indexed adjustment)
    # Actually df.loc[rows] is the last one written. 
    # Pandas writes header at startrow, data at startrow+1.
    # So if startrow=3, header is at 3, data starts at 4.
    # 30 rows of data take 4 to 33. Totals row is at 34.
    
    # Explicitly overwrite the last row with bold format
    # The 'to_excel' writes the data, we just overlay style if needed or write explicitly.
    # To be safe and simple, let's just let pandas write it and we rely on border settings?
    # Actually, xlsxwriter conditional formatting or row-writing is better.
    # Let's just overwrite the "TOTALS" cell label to be safe.
    last_row = 3 + rows + 1
    worksheet.write(last_row, 0, "TOTALS", total_format)
    for col in range(1, 5):
        worksheet.write(last_row, col, "____", total_format)

    # Title & Instructions
    worksheet.merge_range('A1:E1', 'TRASH ROTATION TRACKER (v2)', title_format)
    worksheet.merge_range('A2:E2', 'FLOW: KM -> NP -> PS -> LS -> (Loop)', instr_format)

    # Footer Instructions
    footer_row = last_row + 2
    worksheet.merge_range(f'A{footer_row}:E{footer_row}', 
                          "THE DENT METHOD: Find first empty box. Press House Key firmly into box [     ] to mark done.", 
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
    generate_tracker_v2()
