"""
make_gamesheets.py
~~~~~~~~~~~~~~~~~~

Read in the CSV in '../data/season2015.csv' and
write out excel files with the whole season of game sheets.

If you want to use your own formatting, the documentation for
openpyxl is here: http://openpyxl.readthedocs.org/
"""
from __future__ import print_function
import os

import openpyxl
import pandas as pd


# Change the path if your file is not here.
input_file = os.path.join('data', 'nfl_season2015.csv')
output_file = os.path.join('excel_files', 'sheets2015.xlsx')

# Default for Pandas is to draw a border around the header columns.
# Explicitly turn that off.
pd.core.format.header_style = None

# Fill colors for conditional formatting
def make_fill(color):
    try:
        return openpyxl.styles.PatternFill(
                start_color=color, end_color=color, fill_type='solid')
    except AttributeError:
        fill = openpyxl.styles.Fill()
        fill.fill_type = openpyxl.styles.Fill.FILL_SOLID
        fill.start_color = color
        fill.end_color = color
        return fill

fills = dict(
        red = make_fill('FFFFCCCC'),
        yellow = make_fill('FFFFFFDD'),
        green = make_fill('FFCCFFCC')
)

def add_conditional_fill(sheet, cell_range, color='yellow', formula=None):
    try:
        sheet.conditional_formatting.add(cell_range, {
                'type': 'expression',
                'dxf': {'fill': fills[color]},
                'formula': [formula],
                'stopIfTrue': '1'
        })
    except AttributeError:
        sheet.conditional_formatting.addCustomRule(cell_range, {
                'type': 'expression',
                'dxf': {'fill': fills[color]},
                'formula': [formula],
                'stopIfTrue': '1'
        })

# rows and columns to skip
skip_rows = 7
skip_cols = 1

# Change the path if your file is not here.
print("Reading schedule from {}...".format(input_file))
season = pd.read_csv(input_file)
# Add blank columns for convenience
season['winhome'] = ''
season['spacer'] = ''
season['winaway'] = ''

# Open an excel workbook and write to it.
with pd.ExcelWriter(output_file, engine='openpyxl') as workbook:
    for week, games in season.groupby('week'):
        gametable = games[['dayofweek', 'homeTeam', 'winhome', 'spacer', 'awayTeam', 'winaway']]
        nrow, ncol = gametable.shape
        sheet_name='Week {}'.format(week)
        gametable.to_excel(workbook,
            sheet_name=sheet_name,
            index=False,  # don't show the row numbers
            header=['Day', 'Home', 'Win', '', 'Away', 'Win'],  # alternative column names
            startrow=skip_rows,
            startcol=skip_cols)
        sheet = workbook.sheets[sheet_name]
        # Add instructions
        sheet['A1'] = "Week {} (Sunday is {})".format(week, games[games['dayofweek']=='Sunday']['date'].iloc[0])
        sheet['A2'] = ("Mark the 'Win' column for each team you think "
                       "will win. Winning sheet has the most correct.")
        sheet['A4'] = "Name:"
        sheet['B4'] = "<<your name>>"
        sheet['A5'] = "Tiebreaker *:"
        sheet['B5'] = "<<winning guess is closest to the total combined points in the final game>>"
        sheet['A{}'.format(nrow + skip_rows + 1)] = "Tiebreaker game *"
        # Set column widths
        col_widths = zip('ABCDEFG', (18, 9, 20, 5, 10, 20, 5))
        for col, width in col_widths:
            sheet.column_dimensions[col].width = width
            # Bold and center the headers
            cell = col + str(skip_rows + 1)
            try:
                sheet[cell].font = openpyxl.styles.Font(bold=True)
                sheet[cell].alignment = openpyxl.styles.Alignment(horizontal = 'center')
            except TypeError:
                sheet[cell].style.font = openpyxl.styles.fonts.Font()
                sheet[cell].style.font_style = openpyxl.styles.fonts.Font.bold = True
                sheet[cell].style.alignment = openpyxl.styles.alignment.Alignment()
                sheet[cell].style.alignment_style = openpyxl.styles.alignment.Alignment.HORIZONTAL_CENTER
        # Right align specific cells
        right_align = ('A4', 'A5', 'A{}'.format(nrow + skip_rows + 1))
        for cell in right_align:
            try:
                sheet[cell].alignment = openpyxl.styles.Alignment(horizontal = 'right')
            except TypeError:
                sheet[cell].style.alignment = openpyxl.styles.alignment.Alignment()
                sheet[cell].style.alignment_style = openpyxl.styles.alignment.Alignment.HORIZONTAL_RIGHT
        # Create formatting rules to ensure sheet is correctly filled
        homecol, awaycol = 'D', 'G'
        for row in range(skip_rows + 2, skip_rows + 2 + nrow):
            cell_range = "{startcol}{row}:{endcol}{row}".format(
                    startcol=chr(ord('A') + skip_cols),
                    endcol=chr(ord('A') + skip_cols + ncol - 1),
                    row=row)
            # Yellow if unfilled
            add_conditional_fill(sheet, cell_range, color='yellow', formula=
                    'and(ISBLANK(${homecol}${row}), ISBLANK(${awaycol}${row}))'.format(
                    homecol=homecol, awaycol=awaycol, row=row)
            )
            # Red if both teams picked
            add_conditional_fill(sheet, cell_range, color='red', formula=
                    'and(not(ISBLANK(${homecol}${row})), not(ISBLANK(${awaycol}${row})))'.format(
                    homecol=homecol, awaycol=awaycol, row=row)
            )
            # Green if just one team picked
            add_conditional_fill(sheet, cell_range, color='green', formula=(
                    'or(and(not(ISBLANK(${homecol}${row})), ISBLANK(${awaycol}${row})), '
                    'and(ISBLANK(${homecol}${row}), not(ISBLANK(${awaycol}${row}))))' ).format(
                    homecol=homecol, awaycol=awaycol, row=row)
            )

# finished.
print("Finished.")
print("Output file is at {}.\n".format(output_file))
