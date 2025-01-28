"""
Description:
    Parses a table (list of lists) where each row may have multiple columns.
    - Cells starting with uppercase or numeric dotted references are Section Headers.
    - Cells starting with lowercase dotted references or '-' are treated as List Items.
    - Other cells are treated as Content lines.
    Produces a structured text output grouping content under each Section Header.
"""
#%%
import pandas as pd
import re

def read_excel_file_to_table(filepath, sheet_name=0):
    df = pd.read_excel(filepath, sheet_name=sheet_name, header=None, skiprows=1)
    # Drop entirely empty columns
    df.dropna(how='all', axis=1, inplace=True)
    # Drop empty cells per row
    df = df.apply(lambda x: x.dropna(), axis=1)
    return df

def classify_item(item):
    if pd.isna(item):
        return None
    if isinstance(item, str):
        text = item.strip()
        if not text:
            return None
        # Main header (e.g., "1. ")
        if re.match(r'^\d+\.\s', text):
            return 'Main Header'
        # Multiple levels of subsection headers (e.g., "1.2 ", "1.2.3 ", "1.2.3.4 ")
        elif re.match(r'^\d+(\.\d+)+\s', text):
            return 'Subsection Header'
        elif re.match(r'^[A-Z]\.\s', text) or re.match(r'^[IVXLCDM]+\.\s', text):
            return 'Header'
        elif re.match(r'^[a-z]\.\s', text) or re.match(r'^[ivxlcdm]+\.\s', text):
            return 'List Item'
        else:
            return 'Text'
    else:
        return 'Text'

def df_to_markdown(df, classifications):
    md_lines = []
    for r in range(len(df)):
        for c in range(len(df.columns)):
            if r not in df.index or c not in df.columns:
                continue
            cell_val = df.iloc[r, c]
            cell_class = classifications.iloc[r, c]
            if pd.isna(cell_val) or cell_class is None:
                continue

            text = str(cell_val).strip()

            if cell_class == 'Main Header':
                # h2
                md_lines.append(f"## {text}")
            elif cell_class == 'Subsection Header':
                # Count dots to determine level
                dot_count = text.count('.')  
                # For each dot beyond the first, increase header level up to h6
                # Main header = h2 => one dot is h3 => two dots is h4, etc.
                level = min(2 + dot_count, 6)
                md_lines.append(f"{'#' * level} {text}")
            elif cell_class == 'Header':
                # h2 for standard headers
                md_lines.append(f"## {text}")
            elif cell_class == 'List Item':
                md_lines.append(f"- {text}")
            else:
                md_lines.append(text)
    return "\n".join(md_lines)

file = read_excel_file_to_table('AxactorThesis.xlsx')
file_classifications = file.applymap(classify_item)

markdown_text = df_to_markdown(file, file_classifications)

with open('output.md', 'w') as f:
    f.write(markdown_text)
