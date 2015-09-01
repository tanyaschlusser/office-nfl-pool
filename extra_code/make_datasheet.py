"""
make_datasheet.py
~~~~~~~~~~~~~~~~~~

A one-sheet log for the 2015 season.
Read in the CSV in '../data/nfl_season2015.csv' and
write out to '../excel_files/season2015_datasheet.xlsx'


The input file '../data/nfl_season2015.csv' looks like:

Season,Category,Week,Team,Opponent,AtHome,Points,PointsAllowed,Date,Stadium,Overtime
2008.0,regular,1.0,Arizona Cardinals,San Francisco 49ers,False,23.0,13.0,2008-09-07


and the output file '../excel_files/season2015_datasheet.xlsx' looks like:

week,date,dayofweek,homeTeam,awayTeam
1,2015-09-10,Thursday,New England Patriots,Pittsburgh Steelers

"""
from __future__ import print_function
import os
import sys

# openpyxl is for Excel formatting
import pandas as pd
from openpyxl import styles


# Change the path if your file is not here.
input_filename = os.path.join('data', 'nfl_season2015.csv')
output_filename = os.path.join('excel_files', 'season2015_datasheet.xlsx')

# Default for Pandas is to draw a border around the header columns.
# The below explicitly turns that off.
# (Otherwise writing to .xlsx breaks for some reason)
pd.core.format.header_style = None

print("Reading the season schedule from {}\n...".format(input_filename))
season = pd.read_csv(input_filename)

# Add blank columns for convenience
season['HomePoints'] = ''
season['AwayPoints'] = ''
season['VegasSpread'] = ''
season['HomeFumbles'] = ''
season['AwayFumbles'] = ''
season['HomePenaltyYards'] = ''
season['AwayPenaltyYards'] = ''

# Reorder the columns
season = season[[
        'week', 'date', 'dayofweek',
        'homeTeam', 'HomePoints', 'awayTeam', 'AwayPoints',
        'VegasSpread',
        'HomeFumbles', 'HomePenaltyYards', 'AwayFumbles', 'AwayPenaltyYards'
]]

# Rename the columns
season.columns = [
        'Week', 'Date', 'Day of Week',
        'Home Team', 'Home Points', 'Away Team', 'Away Points',
        'Vegas Spread',
        'Home Fumbles', 'Home Penalty Yards',
        'Away Fumbles', 'Away Penalty Yards']

sheet_name = 'Season 2015'
skip_rows = 0
skip_cols = 0 

# Open an excel workbook and write to it.
with pd.ExcelWriter(output_filename, engine='openpyxl') as workbook:
    season.to_excel(
            workbook,
            sheet_name=sheet_name,
            index=False,  # don't show the row numbers
            startrow=skip_rows,
            startcol=skip_cols)
    sheet = workbook.sheets[sheet_name]
    # Set column widths. You have to change this if you add columns
    col_widths = zip('ABCDEFGHIJKL', (9, 16, 12, 20, 12, 20, 12, 16, 14, 20, 14, 20))
    for col, width in col_widths:
        sheet.column_dimensions[col].width = width
        # Bold and center the headers
        cell = col + str(skip_rows+1)
        try:
            sheet[cell].font = styles.Font(bold=True)
            sheet[cell].alignment = styles.Alignment(horizontal = 'center')
        except TypeError:
            sheet[cell].style.font = styles.fonts.Font()
            sheet[cell].style.font_style = styles.fonts.Font.bold = True
            sheet[cell].style.alignment = styles.alignment.Alignment()
            sheet[cell].style.alignment_style = styles.alignment.Alignment.HORIZONTAL_CENTER
    # Make a horizontal boundary between each week
    try:
        separator = styles.borders.Border(bottom=styles.borders.Side(style='thin'))
    except AttributeError:
        separator = styles.borders.Border()
        separator.border_style = styles.borders.Border.BORDER_THIN
    for row_offset in season[['Week','Home Team']].groupby('Week').count().cumsum().values:
        for col_offset in range(season.shape[1]):
            try:
                sheet.cell(
                    row=skip_rows+row_offset[0]+1,
                    column=skip_cols+col_offset+1
                ).border = separator
            except AttributeError:
                sheet.cell(
                    row=skip_rows+row_offset[0]+1,
                    column=skip_cols+col_offset+1
                ).style.border = separator


# Finished.
print("Finished.")
print("The blank datasheet is in {}\n".format(output_filename))
