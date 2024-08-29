# This files reads the file WordFeud.xlsx and gets all matches from the sheets
# and writes them to the file matches.txt

import openpyxl
import time

def fix_name(name):
    # remove spaces
    name = name.replace(" ", "")

    # remove everything after the second capital letter
    for i in range(2, len(name)):
        if name[i].isupper():
            name = name[:(i + 1)]
            break

    if name == "PelleÅ":
        name = "Pelle"

    if name == "Jöran":
        name = "Joran"

    if name == "JohanL":
        name = "Johan"

    return name

with open("wordfeudligan.txt", "w") as f:
    f.write("game_name Wordfeudligan\n")

    # Open the file
    print("Opening file")
    workbook = openpyxl.load_workbook("WordFeud.xlsx")

    matches = []

    # Loop through all sheets
    print("Looping through sheets")
    for sheet in workbook.worksheets:
        current_matches = []
        # For each row with a value in column F and column G, record the values of column B, C, F and G
        for row in sheet.iter_rows(min_row=2, min_col=2, max_col=7):
            if sheet.title.startswith("Säsong") and row[4].value and row[5].value:
                match = {
                    "sheet": sheet.title,
                    "player1": fix_name(row[0].value),
                    "player2": fix_name(row[1].value),
                    "score1": int(row[4].value),
                    "score2": int(row[5].value)
                }
                current_matches.append(match)
            elif sheet.title.startswith("Cup") and row[2].value and row[3].value:
                match = {
                    "sheet": sheet.title,
                    "player1": fix_name(row[0].value),
                    "player2": fix_name(row[1].value),
                    "score1": int(row[2].value),
                    "score2": int(row[3].value)
                }
                current_matches.append(match)
        # Add matches at the start of the list since sheets are listed in reverse order
        matches = current_matches + matches

    # Write the matches to the file
    # convert start date to seconds since epoch
    start_time = time.mktime(time.strptime("2023-01-01", "%Y-%m-%d"))

    print("Writing to file")
    for match in matches:
        # Convert seconds since epoch to date
        date = time.strftime("%Y-%m-%d %H:%M", time.localtime(start_time))
        f.write("game " + date + "\n")
        f.write(match["player1"] + "\t\t" + str(match["score1"]) + "\n")
        f.write(match["player2"] + "\t\t" + str(match["score2"]) + "\n")
        f.write("\n")

        # Increase the date by 1 day
        start_time += 24 * 60 * 60