from openpyxl.styles import Border, Side, Font, Alignment

class Team:
    def __init__(self, number, name):
        self.number = number
        self.name = name
        self.matches = []
    
    def generate_matches(self, input_matches):
        self.matches = [match for match in input_matches if self in match["red_alliance"] or self in match["blue_alliance"]]

    def generate_sheet(self, current_sheet, provided_by=None):
        thick = Side(border_style="thick")
        thin = Side(border_style="thin")

        current_sheet.merge_cells("B2:E2")
        current_sheet["B2"].value = "Provided by {}".format(provided_by if provided_by != None else "Quantum Robotics")
        current_sheet["B2"].border = Border(top=thick, left=thick, right=thick, bottom=Side(border_style="thin", color="ffffff"))
        current_sheet["B2"].font = Font(sz=14)
        current_sheet["B2"].alignment = Alignment(horizontal="center", vertical="center")

        current_sheet.merge_cells("B3:E3")
        current_sheet["B3"].value = "Team #" + str(self.number) + " | " + self.name
        current_sheet["B3"].border = Border(top=Side(border_style="thin", color="ffffff"), left=thick, right=thick, bottom=thick)
        current_sheet["B3"].font = Font(sz=24)
        current_sheet["B3"].alignment = Alignment(horizontal="center", vertical="center")

        for column, column_name in zip(("B", "C", "D", "E"), ("Match", "Partner", "Opponent 1", "Opponent 2")):
            current_sheet.column_dimensions[column].width = 25

            current_sheet[column + "4"].value = column_name
            current_sheet[column + "4"].border = Border(top=thick, left=thin, right=thin, bottom=thin)
            current_sheet[column + "4"].font = Font(sz=14)
            current_sheet[column + "4"].alignment = Alignment(horizontal="center", vertical="center")

        for index, match in enumerate(self.matches):
            is_red = self in match["red_alliance"]

            current_cell = current_sheet["B"+ str(index+5)]
            current_cell.value = match["name"] + ("\nRed" if is_red else "\nBlue")
            current_cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
            current_cell.font  = Font(sz=14, color="ff0000" if is_red else "0000ff")
            current_cell.alignment = Alignment(horizontal="center", vertical="center")

            partner = next(team for team in (match["red_alliance"] if is_red else match["blue_alliance"]) if team != self)
            current_cell = current_sheet["C"+ str(index+5)]
            current_cell.value = str(partner.number) + "\n" + partner.name
            current_cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
            current_cell.font  = Font(sz=14, color="ff0000" if self in match["red_alliance"] else "0000ff")
            current_cell.alignment = Alignment(horizontal="center", vertical="center")

            for column, opponent in zip(["D", "E"], match["blue_alliance"] if is_red else match["red_alliance"]):
                current_cell = current_sheet[column + str(index+5)]
                current_cell.value = str(opponent.number) + "\n" + opponent.name
                current_cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                current_cell.font  = Font(sz=14, color="ff0000" if self in match["red_alliance"] else "0000ff")
                current_cell.alignment = Alignment(horizontal="center", vertical="center")