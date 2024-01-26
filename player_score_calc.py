import streamlit as st 
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import re
from io import BytesIO
from openpyxl.styles import Alignment, Font
from openpyxl.utils import range_boundaries

DEBUG = True

teams_file = st.file_uploader("upload teams excel file.")

process_button = st.button("process")

players_sheet_name  = "PlayersList"
sheet_to_use = "Copy of Teams"

C_COLOR = "FF00FF00"
VC_COLOR = "FFFFFF00" 


def player_credit_from_excel_sheet(sheet):
    players_credits = {}
    for row in sheet.iter_rows(min_row=2, values_only=True):  # Assuming first row is header
        name, credit = row
        players_credits[name] = credit if credit is not None else 0
    return players_credits

def check_all_team_marked_c_and_vc(sheet):
    last_column_with_data = 0 
    for column in sheet.iter_rows(min_row=1, max_row=2):
        for cell in column:
            if cell.value is not None:
                last_column_with_data = cell.column

    last_col_name = get_column_letter(last_column_with_data)
    
    last_column_with_data = last_column_with_data - 1

    last_row_index = 12
    last_idx = f"{last_col_name}{last_row_index}"

    # # Create the range string
    gph_idx = f"B2:{last_idx}"

    min_col, min_row, max_col, max_row = range_boundaries(gph_idx)
    c_vc_missing = []
    for col in range(min_col, max_col + 1):
        _c, _vc = False, False
        for row in range(min_row, max_row + 1):
            cell = sheet.cell(row=row, column=col)

            clr = cell.fill.start_color.index

            if clr == C_COLOR:
                _c = True
            if clr == VC_COLOR:
                _vc = True
        
        if _c and _vc:
            pass
        else:
            c_vc_missing.append(col)
    
    if len(c_vc_missing) == 0:
        return True, c_vc_missing
    else:
        return False, c_vc_missing

def compute_and_download(excel_data, is_player_sheet_exists):

    player_credit = {}
    for key in st.session_state:
    # Check if the key belongs to your input fields
        if key.startswith("input_"):
            # Access the value of each input field
            value = st.session_state[key]
            # Process the value as needed
            pname = re.sub("input_[\d]+_", "", key)

            player_credit[pname] = eval(value) 
    
    #st.write(player_credit)

    credits_rows = []
    for r in excel_data:
        crow = []
        for r_val in r:
            (pname, factor), = r_val.items()
            crow.append(player_credit.get(pname,0) * factor)
        credits_rows.append(crow)

        
    # for r in excel_data:
    #     crow = []
    #     for pname in r:
    #         crow.append(player_credit.get(pname,0))
    #     credits_rows.append(crow)
    
    # Calculate the number of columns
    num_columns = len(credits_rows[0])

    # Initialize a list with zeros for storing column sums
    column_sums = [0] * num_columns

    column_empty = [""] * num_columns

    # Iterate through each row and column to calculate column sums
    for row in credits_rows:
        for i in range(num_columns):
            column_sums[i] += row[i]

    # Append the column sums to the original 2D list
            
    credits_rows.append(column_empty)
    credits_rows.append(column_sums)

    file_name = teams_file.name

    wb = load_workbook(teams_file, read_only=False)
    
    sheet_names = wb.sheetnames
    sheet_name_lower = [s.lower() for s in sheet_names]

    if players_sheet_name.lower() in sheet_name_lower:
        player_sheet = wb[players_sheet_name]
        wb.remove(player_sheet)
        player_sheet = wb.create_sheet(players_sheet_name)
    else:
        player_sheet = wb.create_sheet(players_sheet_name)
    
    #append the new / updated values
    player_sheet.append(["Name", "Credit"])
    for name, credit in player_credit.items():
        player_sheet.append([name, credit])
    
    # Set the column widths for better readability
    for column in player_sheet.columns:
        max_length = max(len(str(cell.value)) for cell in column)
        player_sheet.column_dimensions[get_column_letter(column[0].column)].width = max_length
    

    if sheet_to_use in wb.sheetnames:
        sheet = wb[sheet_to_use]

        start_row = 60
        start_rol = 2

        for row_index, row in enumerate(credits_rows, start=start_row):  # Start=1 since Excel rows are 1-indexed
            for col_index, value in enumerate(row, start=start_rol):  # Start=1 for columns as well
                cell = sheet.cell(row=row_index, column=col_index)
                cell.alignment = Alignment(horizontal='center')
                cell.font = Font(bold=True)

                cell.value = value
        
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    # Step 4: Create a download button
    btn = st.download_button(
        label="Download updated Excel file",
        data=output,
        file_name="updated_file_"+file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if process_button:
    wb = load_workbook(teams_file, read_only=False)
    
    is_all_okay, missing_team = check_all_team_marked_c_and_vc(wb[sheet_to_use])
    if is_all_okay:

        sheet_names = wb.sheetnames
        sheet_name_lower = [s.lower() for s in sheet_names]
        
        data = []
        if sheet_to_use in wb.sheetnames:
            sheet = wb[sheet_to_use]
            last_column_with_data = 0 
            for column in sheet.iter_rows(min_row=1, max_row=2):
                for cell in column:
                    if cell.value is not None:
                        last_column_with_data = cell.column

            last_col_name = get_column_letter(last_column_with_data)
            
            last_column_with_data = last_column_with_data - 1

            st.write("team count : ", last_column_with_data)

            last_row_index = 12
            last_idx = f"{last_col_name}{last_row_index}"

            # Create the range string
            gph_idx = f"B2:{last_idx}"

            for row in sheet[gph_idx]:
                r_values = []
                for cell in row:
                    clr = cell.fill.start_color.index
                    if clr == C_COLOR:
                        r_values.append({cell.value:2})
                    elif clr == VC_COLOR:
                        r_values.append({cell.value:1.5})
                    else:
                        r_values.append({cell.value:1})
                data.append(r_values)

        if players_sheet_name.lower() in sheet_name_lower:
            st.write("Existing Credit Sheet Found.")
            sheet = wb[players_sheet_name]
            player_credit = player_credit_from_excel_sheet(sheet)

            form = st.form("my_form")
            with form:
                st.header("Player Credits")
                i = 1 
                for name in player_credit:
                    input_key = f"input_{i}_{name}"
                    player_credit[name] = st.text_input(f"{i}.{name}", 
                                                                value=player_credit[name],
                                                                key=input_key)
                    i = i + 1
                submit_btn = form.form_submit_button("Submit", on_click=compute_and_download,args=(data,True) )

        else:
            st.write("No Existing Credit Sheet Found.")
            unique_players_dict = {x for l in data for x in l}

            unique_players_list = list(unique_players_dict)

            form = st.form("my_form")

            wb.close()

            with form:
                i = 1
                input_values = {}
                for name in unique_players_list:
                    input_key = f"input_{i}_{name}"
                    input_values[input_key] = st.text_input(str(i) + "." + name, key=input_key)
                    i = i + 1

                submit_btn = form.form_submit_button("Submit", on_click=compute_and_download,args=(data,False) )
        
    else:
        st.write("All Teams Must Have C and VC")
        st.write(missing_team)


            
                
                    
