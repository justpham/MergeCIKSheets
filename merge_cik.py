"""
Author: Justin Pham
Date: 2/21/2023
Description: Merges the two spreadsheets to match the respective companies' CIK number
"""

import openpyxl
from fuzzywuzzy import fuzz
from fuzzywuzzy import process


def reformat_name(name):
    """
    Removes undesired phrases in a string; also replaces . with a space
    """
    undesired = ["inc.", "limited", "inc", "llc", "llc.", "l.l.c.", "(tiso)", "corp", "corp.", "ltd", "ltd.", "and other issuers", "et al.", "tiso", "l.p.", "lp", "company", "corp", "corporation"]

    # Make all characters lower case and split them into different lists
    temp = name.lower()
    temp = temp.strip(" ")
    temp = temp.split(" ")
    reform_name = ""  # Return value

    # Remove any undesired phrases in the list
    for i in range(len(temp)):
        if temp[i] not in undesired:
            reform_name += temp[i]
            reform_name += " "

    # Change characters to be more readable to the search
    reform_name = reform_name.replace(",", "")
    reform_name = reform_name.replace(".com", " com")
    reform_name = reform_name.rstrip(" .s")
    return reform_name


def compare_name(name1, name2):
    """
    Compares the two company names using fuzzywuzzy name matching
    """

    # Reformat the name so comparing is consistent between two sheets
    company1 = reformat_name(name1)
    company2 = reformat_name(name2)

    # Use fuzzywuzzy name matching to compare the score between the two companies
    ratio = fuzz.ratio(company1, company2)
    return ratio


def binary_company_search(company, current):
    """
    Uses binary search to find a matching company's CIK number
    """

    lo = 2
    hi = 1020564
    while lo <= hi:
        mid = int((lo + hi) / 2)
        # Used to find alphabetical order
        sort = [company, str(past_cik['B'+str(mid)].value)]
        sort.sort(key=str.lower)

        # If the company is found with a 90% match return it
        if compare_name(company, str(past_cik['B'+str(mid)].value)) >= 90:
            new_cik['F'+str(current)].value = past_cik['B'+str(mid)].value
            return past_cik['A'+str(mid)].value
        # If the company is alphabetically lower than the midpoint
        elif sort[0] == company:
            hi = mid - 1
        # If the company is alphabetically higher than the midpoint
        else:
            lo = mid + 1

    # If no match is found
    new_cik['F' + str(current)].value = "N/A"
    return "N/A"


# Open Excel Files for Merge
file1 = openpyxl.load_workbook("InvestigationsJul20toSep21.xlsx")
new_cik = file1['Sheet1']
print("Loaded Workbook 1")
file2 = openpyxl.load_workbook("sec_cik_header_file.xlsx")
past_cik = file2['Sheet1']
print("Loaded Workbook 2")
# Create a new column for "matched company" to help validate results
new_cik['F1'] = "Matched Company"

# Loop for every company name
for x in range(2, 1164):

    print("Searching for " + str(new_cik['A'+str(x)].value) + " (Company #" + str(x - 1) + ")")
    # Search for company in the other spreadsheet using binary search
    cik = binary_company_search(new_cik['A'+str(x)].value, x)
    # Input CIK number if found in the new spreadsheet
    new_cik['E'+str(x)].value = cik
    if cik == "N/A":
        print("No matching CIK found")
    else:
        print("Matching CIK: " + str(cik))

file1.save('MergedCIKsForInvestigationsJul20toSep21.xlsx')
file1.close()
file2.close()
