from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Font, Alignment

input_teams = load_workbook(filename="teams.xlsx")["Sheet1"]
input_matches = load_workbook(filename="matches.xlsx")["Sheet1"]

teams = {}
for i in range(len(input_teams["A"])):
    teams[int(input_teams["A"][i].value)] = {"Name": input_teams["B"][i].value, "Color": [], "Match": [], "Partner": [], "Opponent 1": [], "Opponent 2": []}

for j in range(len(input_matches["A"])):
    for team in teams:
        if team == int(input_matches["B"][j].value):
            teams[team]["Color"].append("Red")
            teams[team]["Match"].append(input_matches["A"][j].value + "\nRed")
            teams[team]["Partner"].append(str(int(input_matches["C"][j].value)) + "\n" + teams[int(input_matches["C"][j].value)]["Name"])
            teams[team]["Opponent 1"].append(str(int(input_matches["D"][j].value)) + "\n" + teams[int(input_matches["D"][j].value)]["Name"])
            teams[team]["Opponent 2"].append(str(int(input_matches["E"][j].value)) + "\n" + teams[int(input_matches["E"][j].value)]["Name"])
        elif team == int(input_matches["C"][j].value):
            teams[team]["Color"].append("Red")
            teams[team]["Match"].append(input_matches["A"][j].value + "\nRed")
            teams[team]["Partner"].append(str(int(input_matches["B"][j].value)) + "\n" + teams[int(input_matches["B"][j].value)]["Name"])
            teams[team]["Opponent 1"].append(str(int(input_matches["D"][j].value)) + "\n" + teams[int(input_matches["D"][j].value)]["Name"])
            teams[team]["Opponent 2"].append(str(int(input_matches["E"][j].value)) + "\n" + teams[int(input_matches["E"][j].value)]["Name"])
        elif team == int(input_matches["D"][j].value):
            teams[team]["Color"].append("Blue")
            teams[team]["Match"].append(input_matches["A"][j].value + "\nBlue")
            teams[team]["Partner"].append(str(int(input_matches["E"][j].value)) + "\n" + teams[int(input_matches["E"][j].value)]["Name"])
            teams[team]["Opponent 1"].append(str(int(input_matches["B"][j].value)) + "\n" + teams[int(input_matches["B"][j].value)]["Name"])
            teams[team]["Opponent 2"].append(str(int(input_matches["C"][j].value)) + "\n" + teams[int(input_matches["C"][j].value)]["Name"])
        elif team == int(input_matches["E"][j].value):
            teams[team]["Color"].append("Blue")
            teams[team]["Match"].append(input_matches["A"][j].value + "\nBlue")
            teams[team]["Partner"].append(str(int(input_matches["D"][j].value)) + "\n" + teams[int(input_matches["D"][j].value)]["Name"])
            teams[team]["Opponent 1"].append(str(int(input_matches["B"][j].value)) + "\n" + teams[int(input_matches["B"][j].value)]["Name"])
            teams[team]["Opponent 2"].append(str(int(input_matches["C"][j].value)) + "\n" + teams[int(input_matches["C"][j].value)]["Name"])

output_wb = Workbook()

for team in teams:
    if output_wb.active.title == "Sheet":
        ws = output_wb.active
        ws.title = str(team)
    else:
        ws = output_wb.create_sheet(title=str(team))

    for column in ("B", "C", "D", "E"):
        ws.column_dimensions[column].width = 25

    thick = Side(border_style="thick")
    thin = Side(border_style="thin")

    ws.merge_cells("B2:E2")
    ws["B2"].value = "Provided by Quantum Robotics"
    ws["B2"].border = Border(top=thick, left=thick, right=thick, bottom=Side(border_style="thin", color="ffffff"))
    ws["B2"].font  = Font(sz=14)
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("B3:E3")
    ws["B3"].value = "Team #" + str(team) + " | " + teams[team]["Name"]
    ws["B3"].border = Border(top=Side(border_style="thin", color="ffffff"), left=thick, right=thick, bottom=thick)
    ws["B3"].font  = Font(sz=24)
    ws["B3"].alignment = Alignment(horizontal="center", vertical="center")

    ws["B4"].value = "Match"
    ws["B4"].border = Border(top=thick, left=thin, right=thin, bottom=thin)
    ws["B4"].font = Font(sz=14)
    ws["B4"].alignment = Alignment(horizontal="center", vertical="center")
    for i, match in enumerate(teams[team]["Match"]):
        curr_cell = ws["B"+ str(i+5)]
        curr_cell.value = match
        curr_cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
        curr_cell.font  = Font(sz=14, color="ff0000" if teams[team]["Color"][i] == "Red" else "0000ff")
        curr_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws["C4"].value = "Partner"
    ws["C4"].border = Border(top=thick, left=thin, right=thin, bottom=thin)
    ws["C4"].font = Font(sz=14)
    ws["C4"].alignment = Alignment(horizontal="center", vertical="center")
    for i, partner in enumerate(teams[team]["Partner"]):
        curr_cell = ws["C"+ str(i+5)]
        curr_cell.value = partner
        curr_cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
        curr_cell.font  = Font(sz=14, color="ff0000" if teams[team]["Color"][i] == "Red" else "0000ff")
        curr_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws["D4"].value = "Opponent 1"
    ws["D4"].border = Border(top=thick, left=thin, right=thin, bottom=thin)
    ws["D4"].font = Font(sz=14)
    ws["D4"].alignment = Alignment(horizontal="center", vertical="center")
    for i, opponent in enumerate(teams[team]["Opponent 1"]):
        curr_cell = ws["D"+ str(i+5)]
        curr_cell.value = opponent
        curr_cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
        curr_cell.font  = Font(sz=14, color="ff0000" if teams[team]["Color"][i] != "Red" else "0000ff")
        curr_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws["E4"].value = "Opponent 2"
    ws["E4"].border = Border(top=thick, left=thin, right=thin, bottom=thin)
    ws["E4"].font = Font(sz=14)
    ws["E4"].alignment = Alignment(horizontal="center", vertical="center")
    for i, opponent in enumerate(teams[team]["Opponent 2"]):
        curr_cell = ws["E"+ str(i+5)]
        curr_cell.value = opponent
        curr_cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
        curr_cell.font  = Font(sz=14, color="ff0000" if teams[team]["Color"][i] != "Red" else "0000ff")
        curr_cell.alignment = Alignment(horizontal="center", vertical="center")

output_wb.save("output.xlsx")