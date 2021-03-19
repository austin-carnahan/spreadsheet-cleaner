import pandas as pd

def fields(rules_file):
    # for multi-sheet workbooks, pandas' pd.read_excel(...) method expects a sheet name to be passed.
    # Instead we'll use pd.ExcelFile(...) to read the file and get access to the contained sheet names.
    xl = pd.ExcelFile(rules_file)

    fields = {}  # making a tmp dict to collect outputs for inspection.

    for sheet_name in xl.sheet_names:
        df = pd.read_excel(rules_file, sheet_name=sheet_name, keep_default_na=False)
        allowed_set = set(df["Allowed Values"]) - {""}
        autocorrect_dict = dict(zip(df["Invalid Value"], df["Corrected Value"]))
        autocorrect_dict = {k: v for k, v in autocorrect_dict.items() if v != ""}

        # throwing these in a dict for storage - they could be written to CSV here instead:
        fields[sheet_name] = {
            "allowed_value_set": allowed_set,
            "autocorrect_dict" : autocorrect_dict
        }
  
    return fields

'''
SAMPLE OUTPUT:
**************

{
    'Race': {
        'allowed_value_set': {
            'Asian',
            'Other Race',
            'Black or African American',
            'Native Hawaiian or Other Pacific Islander',
            'American Indian or Alaska Native',
            'White'
        },
        'autocorrect_dict': {

        }
    },
    'Gender': {
        'allowed_value_set': {
            'M',
            'F',
            'Unknown',
            'Other'
        },
        'autocorrect_dict': {
        'male': 'M',
        'm': 'M',
        'U': 'Unknown',
        'u': 'Unknown',
        'female': 'F',
        'Female': 'F',
        'Unk': 'Unknown',
        'Male': 'M',
        'f': 'F',
        'unk': 'Unknown'
        }
    }
}
'''
