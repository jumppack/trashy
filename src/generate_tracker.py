import pandas as pd
import os

def generate_tracker():
    # 1. SETUP THE DATA
    # ---------------------------------------------------------
    # We create 30 rows of empty brackets "[    ]" for the dent method.
    # We also add a "Turn #" column for tracking.
    rows = 30
    data = {
        "Turn #": list(range(1, rows + 1)),
        "KM": ["[    ]"] * rows,
        "NP": ["[    ]"] * rows,
        "PS": ["[    ]"] * rows,
        "LS": ["[    ]"] * rows
    }

    # Create the DataFrame
    df = pd.DataFrame(data)

    # Add a "TOTALS" row at the bottom
    df.loc[rows] = ["TOTALS", "____", "____", "____", "____"]


    # 2. CREATE THE EXCEL FILE WITH FORMATTING
    # ---------------------------------------------------------
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    filename = os.path.join(output_dir, 'Trash_Rotation_Printable.xlsx')
    
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')

    # Write data to Excel, starting at row 3 (leaving room for a big Header)
    df.to_excel(writer, sheet_name='Sheet1', index=False, startrow=3)

    # Get the workbook and worksheet objects to apply formatting
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # 3. DEFINE STYLES
    # ---------------------------------------------------------
    # Main Body Style: Center aligned, Borders, Monospace font
    cell_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,           # Thin border for inside cells
        'font_name': 'Courier New',
        'font_size': 12
    })

    # Header Style (The names KM, NP, etc.): Bold, Thick Bottom Border
    header_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#D7E4BC', # Light green background for visibility
        'border': 2,           # Thick border
        'font_size': 14
    })

    # Title Style
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 18,
        'align': 'center',
        'valign': 'vcenter'
    })

    # Instruction Style
    instr_format = workbook.add_format({
        'italic': True,
        'font_size': 10,
        'align': 'left'
    })

    # 4. APPLY FORMATTING TO COLUMNS AND ROWS
    # ---------------------------------------------------------
    # Apply the cell format to all data columns (A to E)
    # We make columns B, C, D, E wider to fit the brackets comfortably
    worksheet.set_column('A:A', 8, cell_format)   # Narrower for Turn #
    worksheet.set_column('B:E', 18, cell_format)  # Wider for Names

    # Write the Column Headers manually to apply the 'header_format'
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(3, col_num, value, header_format)

    # Add the Big Title at the top
    worksheet.merge_range('A1:E1', 'TRASH ROTATION TRACKER', title_format)

    # Add the Visual Logic instructions
    worksheet.merge_range('A2:E2', 'FLOW: KM -> NP -> PS -> LS -> (Next Row)', instr_format)

    # Add Footer Instructions (below the table)
    footer_row = rows + 5
    worksheet.merge_range(f'A{footer_row}:E{footer_row}', 
                          "INSTRUCTION: Find first empty box. Press Key/Fingernail into box [ ] to mark done.", 
                          instr_format)

    # 5. PRINT SETTINGS (CRITICAL FOR A4)
    # ---------------------------------------------------------
    worksheet.set_paper(9)  # 9 = A4 Paper Size
    worksheet.fit_to_pages(1, 1) # Forces everything to fit on 1 page width and length
    worksheet.center_horizontally()
    worksheet.set_margins(left=0.5, right=0.5, top=0.5, bottom=0.5)

    # Save the file
    writer.close()

    print(f"Success! File '{filename}' created.")

if __name__ == "__main__":
    generate_tracker()
