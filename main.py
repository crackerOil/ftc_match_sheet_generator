import sys, getopt

from openpyxl import Workbook

from excel_loader import ExcelLoader

# event_url = None
# pdf_mode = False

def main():
    loader = ExcelLoader("./excel/teams.xlsx", "./excel/matches.xlsx")
    teams = loader.load_teams()
    matches = loader.load_matches(teams)

    output_wb = Workbook()
    for team in teams:
        team.generate_matches(matches)

        if output_wb.active.title == "Sheet":
            ws = output_wb.active
            ws.title = str(team.number)
        else:
            ws = output_wb.create_sheet(title=str(team.number))

        team.generate_sheet(ws, provided_by="Quantum Robotics")

    output_wb.save("./excel/output.xlsx")


if __name__ == "__main__":
    help = False

    argumentsList = sys.argv[1:]
    
    options = "hu:"
    long_options = ["help", "url=", "pdf"]
    
    try:
        arguments, values = getopt.getopt(argumentsList, options, long_options)
        
        for currentArgument, currentValue in arguments:
            if currentArgument in ("-h", "--help"):
                print("usage: main.py [-h]\n")
                print("optional arguments:")
                print("\t-h, --help\tshow this help message")
                print("\t-u, --url <event url>\tuse an ftc events url instead of xlsx files")
                print("\t    --pdf\tsave 6 copies of each sheet in a pdf file")  
                help = True
            # elif currentArgument in ("-u", "--url") and not help:
            #     print ("Accessing {} ...".format(currentValue))
            #     event_url = currentValue
            # elif currentArgument == "--pdf" and not help:
            #     print ("Enabling pdf mode.")
            #     pdf_mode = True 
    except getopt.error as err:
        print (str(err))
    else: 
        if not help:
            main()