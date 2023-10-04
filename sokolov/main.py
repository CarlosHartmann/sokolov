'''
sokolov: An LLM-experimenting suite for my singular 'they' research. Will include data migration, cleanup, and checking features.

sokolov sokolov sokolov!
sokolov?
That's right.
'''




# standard libraries
import os
import re
import logging

# installed libraries
from num2words import num2words
import openpyxl
from openpyxl.styles import PatternFill, Font, NamedStyle
from openpyxl.utils import get_column_letter

# project resources
from argparse_assets import handle_args


def get_ord(word: str):
	return num2words(word, ordinal=True)


def read_span(text: str) -> tuple:
    return (int(text.split(',')[0][1:]), int(text.split(',')[1][1:-1]))


def translate_annotation(they_type: str) -> str:
    '''Translate the annotation codes into either of the three major categories.'''
    if they_type in ['plural_they', 'collnoun']:
        return "plural"
    else:
        return "singular"


def adapted(prompt: str, body: str, span: str) -> str:
    span = read_span(span)
    they_form = body[span[0]:span[1]]
    if len(re.findall(they_form.lower(), body.lower())) == 1:
        ordinal = ''
    else:
        matches = [elem.span()[0] for elem in re.finditer(they_form.lower(), body.lower())]
        position = matches.index(span[0]) + 1
        ordinal = f' {get_ord(position)}'

    return prompt.format(ordinal, they_form, body)


def add_metric_formulas(results_sheet, data_sheet, human_col_letter, llm_col_letter, data_sheet_max_row):
    categories = ["singular", "plural", "ambiguous"]
    
    for idx, category in enumerate(categories, start=2):
        tp_formula = f'SUMPRODUCT(({data_sheet}!{human_col_letter}$2:{human_col_letter}${data_sheet_max_row}="{category}")*({data_sheet}!{llm_col_letter}$2:{llm_col_letter}${data_sheet_max_row}="{category}"))'
        fp_formula = f'SUMPRODUCT(({data_sheet}!{human_col_letter}$2:{human_col_letter}${data_sheet_max_row}<>"{category}")*({data_sheet}!{llm_col_letter}$2:{llm_col_letter}${data_sheet_max_row}="{category}"))'
        tn_formula = f'SUMPRODUCT(({data_sheet}!{human_col_letter}$2:{human_col_letter}${data_sheet_max_row}<>"{category}")*({data_sheet}!{llm_col_letter}$2:{llm_col_letter}${data_sheet_max_row}<>"{category}"))'
        fn_formula = f'SUMPRODUCT(({data_sheet}!{human_col_letter}$2:{human_col_letter}${data_sheet_max_row}="{category}")*({data_sheet}!{llm_col_letter}$2:{llm_col_letter}${data_sheet_max_row}<>"{category}"))'
        
        results_sheet.cell(row=idx, column=2).value = f'={tp_formula}'
        results_sheet.cell(row=idx, column=3).value = f'={fp_formula}'
        results_sheet.cell(row=idx, column=4).value = f'={tn_formula}'
        results_sheet.cell(row=idx, column=5).value = f'={fn_formula}'
        
        recall_formula = f'={get_column_letter(2)}{idx}/({get_column_letter(2)}{idx}+{get_column_letter(5)}{idx})'
        precision_formula = f'={get_column_letter(2)}{idx}/({get_column_letter(2)}{idx}+{get_column_letter(3)}{idx})'
        accuracy_formula = f'({get_column_letter(2)}{idx}+{get_column_letter(4)}{idx})/({get_column_letter(2)}{idx}+{get_column_letter(3)}{idx}+{get_column_letter(4)}{idx}+{get_column_letter(5)}{idx})'
        
        results_sheet.cell(row=idx, column=6).value = f'={recall_formula}'
        results_sheet.cell(row=idx, column=7).value = f'={precision_formula}'
        results_sheet.cell(row=idx, column=8).value = f'={accuracy_formula}'
    
    for row in results_sheet.iter_rows(min_row=2, max_row=4, min_col=6, max_col=8):
        for cell in row:
            cell.style = "percent_style"


def get_last_row_with_data(sheet, column="A"):
    # Loop from the bottom up until a non-empty cell is found
    for row in range(sheet.max_row, 0, -1):
        if sheet[f"{column}{row}"].value:
            return row
    return None  # If no data found


def process_experiment_file(file: str, args):
    # Load workbook and sheet
    wb = openpyxl.load_workbook(file)
    data_sheet = wb['data']
    last_row = get_last_row_with_data(data_sheet, "A")

    # Remove unannotated comments
    rows_to_delete = []
    for idx, row in enumerate(data_sheet.iter_rows(min_row=2, max_row=last_row, values_only=True), 2):
        annotated_col_idx = [cell.value for cell in data_sheet[1]].index("annotated")
        if row[annotated_col_idx] != "X":
            rows_to_delete.append(idx)

        noise_col_idx = [cell.value for cell in data_sheet[1]].index("noise")
        if row[noise_col_idx] == "X":
            rows_to_delete.append(idx)

        # Remove comments based on filter_length
        if args.length:
            comment_body_col_idx = [cell.value for cell in data_sheet[1]].index("comment_body")
            if len(row[comment_body_col_idx].value) > args.length:
                rows_to_delete.append(idx)

    # delete all rows marked for deletion
    for idx in reversed(rows_to_delete):
        data_sheet.delete_rows(idx)

    last_row = get_last_row_with_data(data_sheet, "A")

    # Add new columns
    last_col = data_sheet.max_column
    headers = ["prompt", "human_annotation", "LLM_annotation", "inter-annotator_agreement"]
    for idx, header in enumerate(headers, start=1):
        data_sheet.cell(row=1, column=last_col + idx, value=header)

    # Fill in 'prompt' and 'human_annotation' column
    for idx, row in enumerate(data_sheet.iter_rows(min_row=2, max_row=last_row, values_only=True), 2):
        comment_body_col_idx = [cell.value for cell in data_sheet[1]].index("comment_body")
        body = row[comment_body_col_idx]

        human_col_idx = [cell.value for cell in data_sheet[1]].index("human_annotation")

        ambig_col_idx = [cell.value for cell in data_sheet[1]].index("ambiguous")
        ambig = row[ambig_col_idx]

        they_type_idx = [cell.value for cell in data_sheet[1]].index("they_type")
        they_type = row[they_type_idx]

        span_col_idx = [cell.value for cell in data_sheet[1]].index("span")
        span = row[span_col_idx]

        prompt_col_idx = [cell.value for cell in data_sheet[1]].index("prompt")

        if data_sheet.cell(row=idx, column=prompt_col_idx+1).value is None:
            data_sheet.cell(row=idx, column=prompt_col_idx+1).value = adapted(args.prompt, body, span)
        
        if data_sheet.cell(row=idx, column=human_col_idx+1).value is None:
            data_sheet.cell(row=idx, column=human_col_idx+1).value = "ambiguous" if ambig == "X" else translate_annotation(they_type)

    # Add results sheet
    if 'results' not in wb.sheetnames:
        results_sheet = wb.create_sheet('results')
    else:
        results_sheet = wb['results']

    categories = ["singular they", "plural they", "ambiguous they"]
    metrics = ["True Positives", "False Positives", "True Negatives", "False Negatives", "Recall", "Precision", "Accuracy"]

    # Headers
    for idx, metric in enumerate(metrics, start=2):
        results_sheet.cell(row=1, column=idx, value=metric)
        if metric in ["Recall", "Precision", "Accuracy"]:
            results_sheet.cell(row=1, column=idx).fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
            results_sheet.cell(row=1, column=idx).font = Font(bold=True)
    # Rows
    for idx, category in enumerate(categories, start=2):
        results_sheet.cell(row=idx, column=1, value=category)

    # Metrics calculation
    data_sheet_max_row = last_row
    human_annotation_col_letter = get_column_letter([cell.value for cell in data_sheet[1]].index("human_annotation") + 1)
    llm_annotation_col_letter = get_column_letter([cell.value for cell in data_sheet[1]].index("LLM_annotation") + 1)
    wb.add_named_style(NamedStyle(name="percent_style", number_format="0.0%"))
    add_metric_formulas(results_sheet, 'data', human_annotation_col_letter, llm_annotation_col_letter, data_sheet_max_row)

    # Save file
    directory = os.path.dirname(file)
    new_filepath = os.path.join(directory, "experiment.xlsx")
    wb.save(new_filepath)





def main():
    logging.basicConfig(level=logging.NOTSET, format='INFO: %(message)s')
    args = handle_args()

    file = args.inputfile

    if args.task == "preparation":
        process_experiment_file(file, args)

    #TODO:
    #-remove unannotated comments (when there is no "X" in the column called "annotated")
    #-remove comments according to filters
    #-add columns "prompt", "human annotation", "LLM annotation", "inter-annotator agreement"
    #-fill in column "prompt" using args.prompt
    #-add extra sheet for experiment evaluation
    
    # write in prompt, modifying it for each comment


if __name__ == "__main__":
    main()