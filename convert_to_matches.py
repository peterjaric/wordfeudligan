# This files reads the file WordFeud.xlsx and gets all matches from the sheets
# and writes them to the file matches.txt

import openpyxl
import time


# Due to a limitation in Rayter lots of characters are not allowed in the name
def replace_invalid_characters(name):
    name = name.replace(" ", "")

    name = name.replace("å", "aa")
    name = name.replace("ä", "ae")
    name = name.replace("ö", "oe")
    name = name.replace("Å", "Aa")
    name = name.replace("Ä", "Ae")
    name = name.replace("Ö", "Oe")

    return name

def fix_name(name):
    name = name.strip()
    name = replace_invalid_characters(name)

    return name

with open("wordfeudligan.txt", "w") as f:
    f.write("game_name Wordfeudligan\n")

    # Open the file
    print("Opening file")
    workbook = openpyxl.load_workbook("WordFeud.xlsx")

    matches = []

    wordfeudname_to_realname = {
        "äggeth": "PelleAa",
        "Äggeth": "PelleAa",
        "Åkermarken": "Lars",
        "Andreas112": "Andreas",
        "Annafian": "Fia",
        "CamillaB71": "Camilla",
        "dinkelidunk": "Tommy",
        "drsadness": "Jonas",
        "Emma-Victoria": "Emma",
        "gurka1495": "Per",
        "johlits": "JohanL",
        "joppelicious": "Johannes",
        "KaffeKaffe": "Fredrik",
        "Kamomillan": "Camilla",
        "kim.stenberg": "Kim",
        "liljekvist02": "MikaelL", # This Mikael has been modified to MikaelL to not collide with Migus-Mikael
        "Migus": "Mikael",
        "ministerkrister": "Kristofer",
        "peterjaric": "Peter",
        "proso": "PelleA",
        "Scrbell": "Joeran",
        "snilser93": "Svante",
        "ulrika.b.carlsson": "Ulrika",
        "uumartin": "Martin",
    }

    # Loop through all sheets
    print("Looping through sheets")
    for sheet in workbook.worksheets:
        realname_to_wordfeudname = {}

        # Only parse sheets that are seasons or cups
        if (sheet.title.startswith("Säsong") or sheet.title.startswith("Cup")):
            # Parse player names and Wordfeud aliases
            if sheet.title == "Säsong 1":
                # The first season doesn't list the players' Wordfeud names
                realname_to_wordfeudname = {
                    "Jonas": "drsadness",
                    "Pelle": "Äggeth",
                    "Per": "gurka1495",
                    "Kristofer": "ministerkrister"
                }
            else:
                # Find the cell with the name "Wordfeudnamn"
                wordfeud_name_column = None
                wordfeud_name_row = None
                for row in sheet.iter_rows(min_row=1, max_row=100, min_col=1, max_col=20):
                    for cell in row:
                        if cell.value == "Wordfeudnamn":
                            wordfeud_name_column = cell.column
                            wordfeud_name_row = cell.row
                            break
                    if wordfeud_name_column:
                        break

                # Connect the player names with their Wordfeud aliases by iterating through the rows under the cell with the name "Wordfeudnamn"
                for row in sheet.iter_rows(min_row=wordfeud_name_row + 1, min_col=wordfeud_name_column - 1, max_col=wordfeud_name_column):
                    if row[0].value:
                        realname_to_wordfeudname[row[0].value.strip()] = row[1].value.strip()
                    else:
                        # Loop until we find an empty cell
                        break

                # Add the reverse mapping
                # for realname, wordfeudname in realname_to_wordfeudname.items():
                #     if wordfeudname not in wordfeudname_to_realname:
                #         wordfeudname_to_realname[wordfeudname] = fix_name(realname)


            # Parse results
            current_matches = []
            # For each row with a value in column F and column G, record the values of column B, C, F and G
            for row in sheet.iter_rows(min_row=2, min_col=2, max_col=7):
                if sheet.title.startswith("Säsong") and row[4].value and row[5].value:
                    match = {
                        "sheet": sheet.title,
                        # "player1": fix_name(row[0].value),
                        # "player2": fix_name(row[1].value),
                        "player1": wordfeudname_to_realname[realname_to_wordfeudname[row[0].value.strip()]],
                        "player2": wordfeudname_to_realname[realname_to_wordfeudname[row[1].value.strip()]],
                        "score1": int(row[4].value),
                        "score2": int(row[5].value)
                    }
                    current_matches.append(match)
                elif sheet.title.startswith("Cup") and row[2].value and row[3].value:
                    match = {
                        "sheet": sheet.title,
                        # "player1": fix_name(row[0].value),
                        # "player2": fix_name(row[1].value),
                        "player1": wordfeudname_to_realname[realname_to_wordfeudname[row[0].value.strip()]],
                        "player2": wordfeudname_to_realname[realname_to_wordfeudname[row[1].value.strip()]],
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
