from get_markdown import *
from prompt import FEW_SHOT_PROMPT
from llm import *

from google.colab import auth
auth.authenticate_user()

import gspread
from google.auth import default
creds, _ = default()

START_ROW_IDX=2 # Count from 0
SHEET_FILE_NAME = 'Report cell classification'
SAMPLE_SHEET_NAME = 'Hoang_FewshotSamples'
TEST_SHEET_NAME = 'Hoang_FewshotTests'
gc = gspread.authorize(creds)
worksheet = gc.open(SHEET_FILE_NAME).worksheet(SAMPLE_SHEET_NAME)
result_worksheet = gc.open(SHEET_FILE_NAME).worksheet(TEST_SHEET_NAME)
model_call = AnswerGenerator()
all_rows = result_worksheet.get_all_values()

for i in range(9,len(all_rows)):
    # check F{i+1} filled
    if len(all_rows[i]) > 5 and all_rows[i][5] != '':
        continue
    # check B C D E have values
    if all_rows[i][4] == '':
        print(f"No value in E{i+1}. BREAK")
        break
    input1_cell_idx = int(all_rows[i][1]) if all_rows[i][1] != '' else 'Not available'
    input2_cell_idx = int(all_rows[i][2]) if all_rows[i][2] != '' else 'Not available'
    input3_cell_idx = int(all_rows[i][3]) if all_rows[i][3] != '' else 'Not available'
    input_table_idx = int(all_rows[i][4])

    if input1_cell_idx != 'Not available':
        input_example_1 = worksheet.acell(f'H{input1_cell_idx}').value
        output_example_1 = worksheet.acell(f"G{input1_cell_idx}").value
    else:
        input_example_1 = 'Not available'
        output_example_1 = 'Not available'
    if input2_cell_idx != 'Not available':
        input_example_2 = worksheet.acell(f'H{input2_cell_idx}').value
        output_example_2 = worksheet.acell(f"G{input2_cell_idx}").value
    else:
        input_example_2 = 'Not available'
        output_example_2 = 'Not available'
    if input3_cell_idx != 'Not available':
        input_example_3 = worksheet.acell(f'H{input3_cell_idx}').value
        output_example_3 = worksheet.acell(f"G{input3_cell_idx}").value
    else:
        input_example_3 = 'Not available'
        output_example_3 = 'Not available'

    # Construct the prompt with the information
    prompt = FEW_SHOT_PROMPT.format(
        input_example_1=input_example_1,
        output_example_1=output_example_1,
        input_example_2= input_example_2,
        output_example_2=output_example_2,
        input_example_3=input_example_3,
        output_example_3=output_example_3,
        input_table=worksheet.acell(f'H{input_table_idx}').value
    )
    # print(prompt)
    try_times=1
    while True:
        try:
            response = model_call.generate(SYSTEM_MESSAGE, prompt)
            # print("Score:")

            pattern = r"```markdown\n(.*\n*?)\n```"
            is_match = re.search(pattern, response, re.DOTALL)
            markdown_content = is_match.group(1)
            # try_times = 5

            report, matrix = Excel_to_markdown().evaluate_markdown_table(markdown_content, worksheet.acell(f'G{input_table_idx}').value)
            break
        except:
            try_times -=1
            if try_times >=0:
                continue
            else:
                print("No markdown found")
                break
    if try_times <0:
        continue
    # print(report)
    # print(matrix)
    start_idx = 0
    letters = [chr(i) for i in range(ord('J'), ord('Z') + 1)]
    result_worksheet.update(f'F{i+1}', [[worksheet.acell(f'H{input_table_idx}').value]])
    result_worksheet.update(f'G{i+1}', [[worksheet.acell(f'G{input_table_idx}').value]])
    result_worksheet.update(f'H{i+1}', [[markdown_content]])
    result_worksheet.update(f'I{i+1}', [[str(matrix)]])
    for k, v in flatten_dict(report).items():
        if 'f1' not in k:
            continue
        print(f'{letters[start_idx]}{i+1}', k, v)
        result_worksheet.update(f'{letters[start_idx]}{i+1}', [[v]])
        start_idx += 1