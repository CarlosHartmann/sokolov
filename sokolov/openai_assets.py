'''
openai_assets: All things related to the openAI API. 
'''


import openpyxl
import requests

from excel_assets import *

api_key ="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"

def moderation_check(worksheet, output_directory):
    # Identify the "comment_body" column
    for column in worksheet.iter_cols(1, worksheet.max_column):
        if column[0].value == "comment_body":
            comment_column = column
            break
    else:
        print("No 'comment_body' column found.")
        return

    # New workbook for unacceptable comments
    wb_unacceptable = openpyxl.Workbook()
    ws_unacceptable = wb_unacceptable.active

    # Check each cell in the "comment_body" column
    rows_to_delete = []
    for cell in comment_column:
        if cell.row == 1:  # Skip header
            continue
        
        comment = cell.value
        if comment is None:
            break
        categories = get_flagged_categories(comment)
        if categories:
            row_values = [c.value for c in worksheet[cell.row]]
            row_values.append(", ".join(categories))  # Append flagged categories
            ws_unacceptable.append(row_values)
            rows_to_delete.append(cell.row)
    
    # Deleting rows from original worksheet
    for row_num in sorted(rows_to_delete, reverse=True):
        worksheet.delete_rows(row_num)

    # Save the new workbook
    wb_unacceptable.save(output_directory)


def get_flagged_categories(comment):
    # Make the API request to OpenAI's moderation API
    response = requests.post(
        'https://api.openai.com/v1/moderations',
        headers={
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {api_key}',
        },
        json={
            'input': comment
        }
    )

    data = response.json()

    # Extract flagged categories
    flagged_categories = []
    for category, flagged in data['results'][0]['categories'].items():
        if flagged:
            flagged_categories.append(category)
    
    return flagged_categories


def get_llm_response(prompt, llm):
    pass


def interpreted(llm_response):
    pass


def complete_statistics(worksheet):
    #-on the results sheet, fill in IAA columns and overall IAA cell
    pass

def conduct_experiment(workbook: str, llm: str):
    '''
    Central function for conducting experiments.
    Requires a workbook that will have the necessary headers for the experiment (prompt, LLM_response, LLM_annotation, human_annotation, inter-annotator-agreement).
    Also requires a "results" worksheet that will be extended with the total IAA and should already have the necessary formulas to display the overall results.
    '''

    # initiate two required worksheets
    data_sheet = workbook['data']
    results_sheet = workbook['results']

    # get required column numbers
    # for the data sheet
    prompt_col = get_column_by_header(data_sheet, 'prompt')
    response_col = get_column_by_header(data_sheet, 'LLM_response')
    annotation_col = get_column_by_header(data_sheet, 'LLM_annotation')
    human_annotation_col = get_column_by_header(data_sheet, 'human_annotation')
    IAA_col = get_column_by_header(data_sheet, 'inter-annotator_agreement')
    # for the results sheet
    results_plural = get_column_by_header(results_sheet, 'plural')
    plural_iaa = results_plural+2
    results_singular = get_column_by_header(results_sheet, 'singular')
    singular_iaa = results_singular+2

    lastrow = get_last_row_with_data(data_sheet)

    for row in range(2, lastrow+1):
        prompt = data_sheet.cell(row=row, column=prompt_col).value # read prompt
        response = get_llm_response(prompt, llm) # send it to LLM
        data_sheet.cell(row=row, column=response_col).value = response
        data_sheet.cell(row=row, column=annotation_col).value = interpreted(response)

        llm = data_sheet.cell(row=row, column=annotation_col).value
        human = data_sheet.cell(row=row, column=human_annotation_col).value
        data_sheet.cell(row=row, column=IAA_col).value = "X" if llm == human else None

    complete_statistics(results_sheet)
    

    pass