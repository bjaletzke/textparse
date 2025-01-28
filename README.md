# TextParse Excel to Markdown

This is a super simple script written for myself and my odd use case.
Using this, text written in excel in multiple columns can be quickly translated to a markdown file and saved.

## Logic

Row/Column (R/C) pairs that start with an uppercase dotted number or letter are section headers, so that '1.', 'A.' or 'I.' are section headers, and e.g. '2.1' is a subsection header
R/C pairs that start with a dotted number or letter are section headers, so that '-', 'a.' or 'i.' are list items
Columns that start with Strings are the content of the section. Each Text row / column should be on a new line

## Input / Output

Logically, we have to traverse the data so that a structure of:

Input --->

| Section | List | Content |Content 2|
|---------|------|---------| |
| 1. Start| | | |
| | start text | | |
| | a.listitem | | |
| | | a.a. text2 | |
| | | a.b. text3 | |
| | b. listitem2 | | |
| | | b.a.text4 | |
||||||
| 2. Title | | | |
| | maintext | | |
| | 2.1 Section Title | | |
| | | text | |
| | | text | |
| | | a. text | |
| | additional main text | | |
| | | subtext main | |

Produces the result:

OUTPUT --->

## 1. Start

start text
a. listitem
a.a. text2
a.b. text3
b. listitem2
b.a. text

## 2. Title

maintext

### 2.1 Section Title

text
text
a. text
additional main text
subtext main
