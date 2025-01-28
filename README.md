# TextParse Excel to Markdown

This is a super simple script written for myself and my odd use case.
Using this, text written in excel in multiple columns can be quickly translated to a markdown file and saved.

## Input / Output

### Content logic
These will be h2 headers:
- Numbered headers (e.g. "1. ", "2. ")
- Uppercase letters (e.g. "A.", "B.")
- Uppercase Roman numerals (e.g. "I.", "II.")
    
These will be h3+ headers depending on depth
- Subsection headers (e.g. "1.1", "2.1.1")

List items identified by:
- Lowercase letters (e.g. "a.", "b.")
- Lowercase Roman numerals (e.g. "i.", "ii.")

Regular text content will be parsed as such

Input Excel:                    Output Markdown:
1. Main                        ## 1. Main
text                          text
a. list                       - list
i. sublist                    1. sublist
ii. sublist                   2. sublist
b. list                       - list
1.1 Subsection               ### 1.1 Subsection
content                      content