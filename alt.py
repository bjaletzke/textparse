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
        # Multiple levels of subsection headers (e.g., "1.2 ", "1.2.3 ")
        elif re.match(r'^\d+(\.\d+)+\s', text):
            return 'Subsection Header'
        # Uppercase letter or roman numeral headers with dots (e.g., "A.A", "III.I")
        elif re.match(r'^[A-Z](\.[A-Z])+\s', text) or re.match(r'^[IVXLCDM]+(\.[IVXLCDM]+)+\s', text):
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
    list_stack = []  # Track nested lists
    current_level = 0
    
    def get_list_marker(text, level):
        # Check if it's roman numerals
        if re.match(r'^[ivxlcdm]+\.\s', text.lower()):
            return f"{len(list_stack) + 1}."
        # Check if it's alphabetic
        elif re.match(r'^[a-z]\.\s', text):
            return f"    " * level + "-"
        # Numbered list
        else:
            return f"{len(list_stack) + 1}."
    
    for r in range(len(df)):
        for c in range(len(df.columns)):
            if r not in df.index or c not in df.columns:
                continue
            
            cell_val = df.iloc[r, c]
            cell_class = classifications.iloc[r, c]
            
            if pd.isna(cell_val) or cell_class is None:
                continue

            text = str(cell_val).strip()
            
            if cell_class in ['Main Header', 'Subsection Header', 'Header']:
                list_stack = []
                current_level = 0
                if cell_class == 'Main Header':
                    md_lines.append(f"\n## {text}")
                elif cell_class == 'Subsection Header':
                    # Count dots for nesting level (A.A or III.I will have 1 dot)
                    dots = text.count('.')
                    level = min(2 + dots, 6)  # markdown supports h1-h6
                    md_lines.append(f"\n{'#' * level} {text}")
                else:
                    md_lines.append(f"\n## {text}")
            
            elif cell_class == 'List Item':
                content = text.split('.', 1)[1].strip()
                marker = get_list_marker(text, current_level)
                md_lines.append(f"{marker} {content}")
                
                if marker.startswith('    '):  # Sub-item
                    current_level += 1
                else:
                    current_level = 0
                list_stack.append(marker)
            
            else:  # Regular text
                md_lines.append(f"{'    ' * current_level}{text}")
    
    return "\n".join(md_lines)

file = read_excel_file_to_table('AxactorThesis.xlsx')
file_classifications = file.applymap(classify_item)

markdown_text = df_to_markdown(file, file_classifications)

with open('output.md', 'w') as f:
    f.write(markdown_text)