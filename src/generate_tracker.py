import pandas as pd
import os

def generate_tracker_final():
    # ---------------------------------------------------------
    # 1. SETUP THE DATA
    # ---------------------------------------------------------
    rows = 50
    dent_bracket = "[     ]" 
    
    data = {
        "Turn #": list(range(1, rows + 1)),
        "KM": [dent_bracket] * rows,
        "NP": [dent_bracket] * rows,
        "PS": [dent_bracket] * rows,
        "LS": [dent_bracket] * rows
    }

    df = pd.DataFrame(data)
    # Add Totals Row
    df.loc[rows] = ["TOTALS", "____", "____", "____", "____"]

    # ---------------------------------------------------------
    # 2. CREATE THE EXCEL FILE
    # ---------------------------------------------------------
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    filename = os.path.join(output_dir, 'Trash_Rotation_Printable.xlsx')
    
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    
    # Write data starting at row 3 (Header is row 3, Data starts row 4)
    # We will write without the default header first to control formatting manually
    df.to_excel(writer, sheet_name='Sheet1', index=False, startrow=3, header=False)

    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    worksheet.hide_gridlines(2) # Hide view lines

    # ---------------------------------------------------------
    # 3. DEFINE STYLES
    # ---------------------------------------------------------
    # Base font
    base_props = {'font_name': 'Courier New', 'font_size': 11, 'align': 'center', 'valign': 'vcenter'}

    # 1. Internal Thin Grid
    format_internal = workbook.add_format({**base_props, 'border': 1})

    # 2. Header Style (Thick Top/Side, Thin Bottom, or Boxed)
    format_header = workbook.add_format({
        **base_props,
        'bold': True,
        'fg_color': '#D7E4BC',
        'font_size': 14,
        'left': 2, 'right': 2, 'top': 2, 'bottom': 1 # Thick outside
    })
    
    # 3. Data Rows (Side Borders Thick, Internal Thin)
    # We need separate formats for Left Column, Middle Columns, Right Column to make the "Outer Border" thick.
    
    # Left Edge Data
    format_left = workbook.add_format({**base_props, 'left': 2, 'right': 1, 'top': 1, 'bottom': 1})
    # Middle Data
    format_mid  = workbook.add_format({**base_props, 'left': 1, 'right': 1, 'top': 1, 'bottom': 1})
    # Right Edge Data
    format_right = workbook.add_format({**base_props, 'left': 1, 'right': 2, 'top': 1, 'bottom': 1})

    # 4. Totals Row Style (Thick Bottom/Side)
    format_total_left = workbook.add_format({
        **base_props, 'bold': True, 'font_size': 12,
        'left': 2, 'right': 1, 'top': 1, 'bottom': 2
    })
    format_total_mid = workbook.add_format({
        **base_props, 'bold': True, 'font_size': 12,
        'left': 1, 'right': 1, 'top': 1, 'bottom': 2
    })
    format_total_right = workbook.add_format({
        **base_props, 'bold': True, 'font_size': 12,
        'left': 1, 'right': 2, 'top': 1, 'bottom': 2
    })

    # Title & Instructions (Text Wrap is CRITICAL)
    title_format = workbook.add_format({
        'bold': True, 'font_size': 18, 'align': 'center', 'valign': 'vcenter',
        'font_name': 'Arial', 'text_wrap': True
    })
    instr_format = workbook.add_format({
        'italic': True, 'font_size': 10, 'align': 'center', 'valign': 'vcenter',
        'font_name': 'Arial', 'text_wrap': True
    })

    # ---------------------------------------------------------
    # 4. APPLY FORMATTING
    # ---------------------------------------------------------
    # Set Column Widths
    worksheet.set_column('A:A', 10)
    worksheet.set_column('B:E', 22)

    # A. Write Header (Row 3) manually
    headers = df.columns.values
    for col, value in enumerate(headers):
        # Determine border style for header
        style = format_header
        # If we want specific thick left/right on the *ends* of the header:
        # Currently format_header has left=2, right=2. 
        # But for middle columns, we usually want thin borders between them? 
        # Let's customize if picky, but user just said "Thick borders around main table".
        # So: Top Row gets Top Thick. Left Col gets Left Thick. Right Col gets Right Thick. Bottom Row gets Bottom Thick.
        
        # Cloning formatting for edges is safer
        p = {**base_props, 'bold': True, 'fg_color': '#D7E4BC', 'font_size': 14, 'top': 2, 'bottom': 1}
        if col == 0:
            p['left'] = 2
            p['right'] = 1
        elif col == len(headers) - 1:
            p['left'] = 1
            p['right'] = 2
        else:
            p['left'] = 1
            p['right'] = 1
        
        f = workbook.add_format(p)
        worksheet.write(3, col, value, f)

    # B. Write Data Rows (Row 4 to 33)
    # Range is 0 to 29 (30 rows).
    start_data_row = 4
    for r in range(rows):
        current_excel_row = start_data_row + r
        # Turn # (Col 0)
        worksheet.write(current_excel_row, 0, data["Turn #"][r], format_left)
        # Person Cols (1, 2) -> Middle
        worksheet.write(current_excel_row, 1, dent_bracket, format_mid)
        worksheet.write(current_excel_row, 2, dent_bracket, format_mid)
        # Person Col (3, 4 is last?) -> 'LS' is index 3 in list (0,1,2,3). Wait, df columns are 5 (Turn, KM, NP, PS, LS)
        # Indices: 0, 1, 2, 3, 4.
        
        # Fix logic:
        # Col 0: Left
        # Col 1, 2, 3: Mid
        # Col 4: Right
        worksheet.write(current_excel_row, 1, dent_bracket, format_mid) # KM
        worksheet.write(current_excel_row, 2, dent_bracket, format_mid) # NP
        worksheet.write(current_excel_row, 3, dent_bracket, format_mid) # PS
        worksheet.write(current_excel_row, 4, dent_bracket, format_right) # LS

    # C. Write Totals Row (Row 34)
    totals_row = start_data_row + rows
    # Col 0
    worksheet.write(totals_row, 0, "TOTALS", format_total_left)
    # Col 1, 2, 3
    worksheet.write(totals_row, 1, "____", format_total_mid)
    worksheet.write(totals_row, 2, "____", format_total_mid)
    worksheet.write(totals_row, 3, "____", format_total_mid)
    # Col 4
    worksheet.write(totals_row, 4, "____", format_total_right)

    # ---------------------------------------------------------
    # 5. TITLE & INSTRUCTIONS (Full Width)
    # ---------------------------------------------------------
    # Set Row Heights for title to ensure visibility
    worksheet.set_row(0, 30) # Title
    worksheet.set_row(1, 20) # Flow
    worksheet.set_row(3, 25) # Header height

    worksheet.merge_range('A1:E1', 'TRASH ROTATION TRACKER', title_format)
    worksheet.merge_range('A2:E2', 'FLOW: KM -> NP -> PS -> LS -> (Loop)', instr_format)

    # Footer Instructions
    footer_row = totals_row + 2
    worksheet.set_row(footer_row, 30) # Give extra space for wrapped text
    worksheet.merge_range(f'A{footer_row+1}:E{footer_row+1}', 
                          "INSTRUCTION: Find first empty box. Press House Key firmly into box [     ] to mark done.", 
                          instr_format)

    # ---------------------------------------------------------
    # 6. PAGE LAYOUT
    # ---------------------------------------------------------
    worksheet.set_paper(9)  # A4
    worksheet.fit_to_pages(1, 1) 
    worksheet.center_horizontally()
    worksheet.set_margins(left=0.5, right=0.5, top=0.5, bottom=0.5)
    
    # Set Print Area to ensure everything is captured
    # Header is A1. Footer is at `footer_row`.
    worksheet.print_area(0, 0, footer_row+1, 4)

    writer.close()
    print(f"Success! File '{filename}' created with refined formatting.")

if __name__ == "__main__":
    generate_tracker_final()
