import pandas as pd
from pandas_schema import Column, Schema
from pandas_schema.validation import InListValidation
import xlsxwriter, os

# return object
class SpreadCheck:
    def __init__(self, spreadsheet):
        self.spreadsheet = spreadsheet
        self.errors = []
        self.messages = []
        self.corrections = 0;
        self.flags = 0;

        
def clean(spreadsheet, rules, tempfolder, progress=None):

    if not os.path.exists(tempfolder):
            os.makedirs(tempfolder)
    filepath = tempfolder +'/temporary_file.xlsx'

    results = SpreadCheck(filepath)

    ################################
    #   LOAD SPREADSHEET TO DF     #
    ################################
    try:
        df = pd.read_csv(spreadsheet)
    except:
        try:
            df = pd.read_excel(spreadsheet)
        except:
            results.spreadsheet = None
            results.errors.append("ERROR: Unaccepted File Format")
            return results


    ################################
    #   SCHEMAS VALIDATION         #
    ################################
    schemas_list = []

    df.fillna("NULL", inplace= True)

    for field, rule in rules.items():
        schemas_list.append(Column(field, [InListValidation(list(rule['allowed_value_set']))]))

    schema = Schema(schemas_list)
    errors = schema.validate(df, columns=schema.get_column_names())

    # reset progress bar
    if progress:
        progress.config(value=0)
    
    # increment progress bar #1
    if progress:
        progress.step(20)


    ################################
    #   BUILD SHEET W FORMATTING   #
    ################################

    writer = pd.ExcelWriter(filepath, engine='xlsxwriter')

    # Skip row 1 headers so we can add manunally with formatting
    df.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False, index=False)

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    ### WORKBOOK FORMATS ###
    yellow_highlight = workbook.add_format({'bg_color': '#FFEB9C' })

    header = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D7E4BC',
        'border': 1
        })

    # increment progress bar #2
    if progress:
        progress.step(20)

    # Set column widths
    worksheet.set_column(0, len(df.columns)-1, 22)
    worksheet.set_default_row(hide_unused_rows=True)

    # Write the column headers with the defined format.
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header)

    # increment progress bar #3
    if progress:
        progress.step(20)

    flag_rows = []
    df_length = len(df)

    for error in errors:
        
        # Catch Key errors, like rules_file, data_file column name mismatch
        try:
            row = error.row + 1
            column = df.columns.get_loc(error.column)
        except:
            results.spreadsheet = None
            results.errors.append("ERROR: " + str(error))
            return results

        # Catch other errors. Not sure what yet.
        try:
            # If an autocorrect mapping exists for value, replace it.
            if error.value in rules[error.column]['autocorrect_dict'].keys():
                worksheet.write(row, column, rules[error.column]['autocorrect_dict'][error.value])
                results.corrections += 1
            else:
                flag_rows.append(error.row)
                # If no autocorrect mapping exists, highlight and annotate the entry
                # Comments
                worksheet.write_comment(row, column , error.message)
                # Highlights
                worksheet.conditional_format(row, column, row, column, {'type': 'no_errors', 'format': yellow_highlight})
                results.flags += 1
        except Exception as err:
            results.spreadsheet = None
            results.errors.append("ERROR: " + err)
            return results
            
    # increment progress bar #4
    if progress:
        progress.step(20)

    # Hide Rows that don't contain annotations
    for i in range(df_length+1):
        if i not in flag_rows:
            worksheet.set_row(i + 1, None, None, {'hidden': True})

    # increment progress bar #5
    if progress:
        progress.step(20)
    
    writer.save()

    results.messages.append("SUCCESS:")
    results.messages.append("{} entries corrected".format(results.corrections))
    results.messages.append("{} entries flagged for review".format(results.flags))
    return results
