import streamlit as st 
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import re
from io import BytesIO
from openpyxl.styles import Alignment, Font

DEBUG = True

teams_file = st.file_uploader("upload teams excel file.")

process_button = st.button("process")

# def transpose(l1, l2):
 
#     # iterate over list l1 to the length of an item 
#     for i in range(len(l1[0])):
#         # print(i)
#         row =[]
#         for item in l1:
#             # appending to new list with values and index positions
#             # i contains index position and item contains values
#             row.append(item[i])
#         l2.append(row)
#     return l2


def compute_and_download(excel_data):

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
        for pname in r:
            crow.append(player_credit.get(pname,0))
        credits_rows.append(crow)
    
    # credit_by_team = []
    # transpose(credits_rows,credit_by_team)

    # #st.write(credit_by_team)

    # for sublist in credit_by_team:
    #     # Calculate the sum of the sublist
    #     sublist_sum = sum(sublist)
    #     # Append the sum to the sublist
    #     sublist.append(sublist_sum)
    

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
    if "Teams" in wb.sheetnames:
        sheet = wb["Teams"]

        start_row = 50
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
    if "Teams" in wb.sheetnames:
        sheet = wb["Teams"]
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

        data = []
        for row in sheet[gph_idx]:
            row_data = [cell.value for cell in row]
            data.append(row_data)
        
        unique_players_dict = {x for l in data for x in l}

        # for row in data:
        #     st.write(row)

        unique_players_list = list(unique_players_dict)

        form = st.form("my_form")

        wb.close()

        with form:
            i = 1
            input_values = {}
            for pname in unique_players_list:
                input_key = f"input_{i}_{pname}"
                input_values[input_key] = st.text_input(str(i) + "." + pname, key=input_key)
                i = i + 1

            submit_btn = form.form_submit_button("Submit", on_click=compute_and_download,args=(data,) )

           
            
                