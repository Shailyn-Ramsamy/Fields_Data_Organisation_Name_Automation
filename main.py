import pandas as pd
from thefuzz import fuzz, process
import openpyxl
from openpyxl.styles import PatternFill
from concurrent.futures import ThreadPoolExecutor


def highlight_cells(val):
    if val in uncertain:
        color = 'yellow'
    elif "To Be Verified" in val:
        color = 'red'
    elif val in uncertain_2:
        color = 'pink'
    else:
        color = ''
    return 'background-color: {}'.format(color)


def handle_special_cases(first_column_value):
    if fuzz.token_sort_ratio(first_column_value, "Save The Children") > 98:
        return "Save The Children International"
    else:
        return None


def find_match(row):
    first_column_value = str(row['Input'])
    special_case = handle_special_cases(first_column_value)
    if special_case:
        return special_case
    max_score = 0
    match = "To Be Verified"
    for second_column_value in unique_org_names:
        # Calculate the similarity score between the two strings
        score = fuzz.token_sort_ratio(first_column_value, second_column_value)
        score2 = fuzz.ratio(first_column_value, second_column_value)
        score3 = fuzz.partial_ratio(first_column_value, second_column_value)

        # Check if the score is above a certain threshold
        if score >= 93 or score2 >= 93 and score > max_score and score2 > max_score:
            # If so, store the value of the second column
            max_score = score
            match = second_column_value

        elif score >= 77 and score < 93 and score2 >= 77 and score2 < 93 and score3 == 100 and score > max_score:
            max_score = score
            uncertain.add(second_column_value)
            match = second_column_value

        elif score >= 50 and score < 77 and score2 >= 50 and score3 == 100 and score2 < 77 and score > max_score:
            max_score = score
            uncertain_2.add(second_column_value)
            match = second_column_value
    return match


df = pd.read_excel(r"Input.xlsx")
df2 = pd.read_excel(r"FD_Orgs.xlsx")

unique_org_names = set(df2.iloc[:, 0])

df['Input'] = df['Input'].astype(str)

mask = df.iloc[:, 0].notnull()

# Create a list to store the matches
matches = []

# Create sets to store uncertain and uncertain_2 values
uncertain = set()
uncertain_2 = set()

with ThreadPoolExecutor() as executor:
    results = executor.map(find_match, df[mask].to_dict('records'))
    matches = list(results)

df['Matches'] = matches

# this highlights cells yellow in the dataframe
styled_df = df.style.applymap(highlight_cells)

display(styled_df)
