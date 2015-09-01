"""
transform
~~~~~~~~~

Helper functions for data manipulation using Pandas.
"""
import numpy as np
import pandas as pd


def from_byteam_to_bygame(df, augment=True, dont_mirror=[]):
    """Tranform data with one row per team to one row per game.

    In the 'byteam' format, there is one row per team -- one for the
    Home team and one for the Away team. In the 'bygame' format, there
    is one row per game, plus one row per team per bye week.

    If 'augment' is True, then prefix every column not in
    'dont_mirror' with an 'A_' for 'Away' or an 'H_' for 'Home'
    to make separate columns for each team.
    """
    common_cols = ['Season', 'Week', 'Home', 'Away']
    home = df[df.AtHome]
    del home['AtHome']
    if not augment:
        return home
    # Otherwise augment the stats so that each column in the
    # dataset is prefixed with an 'H' for the home team or 'A' for
    # the away team.
    away = df[df.AtHome == False]
    del away['AtHome']
    new_home_colnames = []
    new_away_colnames = []
    for c in home.columns:
        if c in common_cols:
            new_home_colnames.append(c)
            new_away_colnames.append(c)
        elif c == 'Team':
            new_home_colnames.append('Home')
            new_away_colnames.append('Away')
        elif c == 'Opponent':
            new_home_colnames.append('Away')
            new_away_colnames.append('Home')
        elif c in dont_mirror:
            new_home_colnames.append(c)
            del away[c]
        else:
            new_home_colnames.append('H_' + c)
            new_away_colnames.append('A_' + c)
    home.columns = new_home_colnames 
    away.columns = new_away_colnames 
    # Left join because the 'bye' weeks are designated 'AtHome'
    return home.merge(away, on=common_cols, how='left')
    
    

def add_derived_columns(df):
    """Add 'Spread', 'WinPct', and a flag 'LastWkBye'

    Make your own here if you like.
    """
    # Spread
    df['Spread'] = df.Points - df.PointsAllowed
    # WinPct
    #   Cumulative win percent to date in a season.
    df['WinPct'] = df.groupby(('Team','Season'))['Spread'].apply(
        lambda s: ( 0.5 * (s == 0).cumsum() + (s > 0).cumsum() )
                / s.notnull().cumsum() )
    # LastWkBye
    #   True if last week's game was a Bye.
    #   This requires grouping by Team and Season.
    #   Be sure to sort first.
    df.sort(['Team', 'Season','Week'], inplace=True)
    df['LastWkBye'] = df.groupby(['Team','Season'])['Spread'
            ].shift(1).isnull().fillna(False)
    #   The first week is null because it was shifted; make it False.
    df.ix[df.Week == 1, 'LastWkBye'] = False


def add_rolling_mean(df, cols, prefix='m{window}wk_', window=5, min_periods=2):
    """Add a rolling mean to each column in 'cols'.

    Prefix the new columns with the stated prefix.
    Sort the data by Team, Season, and Week, and group by Team
    when calculating the moving averages, so the averages
    will cross over seasons.
    """
    df.sort(['Team', 'Season', 'Week'], inplace=True)
    for c in cols:
        colname = prefix.format(window=window) + c
        df[colname] = df.groupby('Team')[c].apply(
                pd.rolling_mean, window=window, min_periods=min_periods)


def add_ewma(df, cols, prefix='ewma{center}wk_', center=2):
    """Add an exponentially weighted moving average to columns in 'cols'.

    Prefix the new columns with the stated prefix.
    Sort the data by Team, Season, and Week, and group by Team
    when calculating the moving averages, so the averages
    will cross over seasons.
    """
    df.sort(['Team', 'Season', 'Week'], inplace=True)
    for c in cols:
        colname = prefix.format(center=center) + c
        df[colname] = df.groupby('Team')[c].apply(
                pd.ewma, center)


def add_lag(df, cols, prefix='lag{lag}_', lag=1):
    """Add the stated lag to each column in 'cols'.

    Prefix the new columns with the stated prefix.
    Sort the data by Team, Season, and Week, and group by Team,
    so the lags will cross over seasons.
    """
    df.sort(['Team', 'Season', 'Week'], inplace=True)
    for c in cols:
        colname = prefix.format(lag=lag) + c
        df[colname] = df.groupby('Team')[c].shift(lag)
