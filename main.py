import sys, getopt

from openpyxl import Workbook

from excel_loader import ExcelLoader
from ftc_api_loader import FTCApiLoader

def main(event_url, api_token, provided_by):
    if event_url == None:
        try:
            loader = ExcelLoader("./excel/teams.xlsx", "./excel/matches.xlsx")
        except:
            print("Missing input files.")
        else:
            teams = loader.load_teams()
            matches = loader.load_matches(teams)
    else:
        loader = FTCApiLoader(event_url, api_token)
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

        team.generate_sheet(ws, provided_by)

    output_wb.save("./excel/output.xlsx")


if __name__ == "__main__":
    event_url = None
    api_user = None
    api_key = None
    # pdf_mode = False
    provided_by = None
    help = False

    argumentsList = sys.argv[1:]
    
    options = "hu:"
    long_options = ["help", "url=", "pdf", "provided-by="]
    
    try:
        arguments, values = getopt.getopt(argumentsList, options, long_options)
        
        for currentArgument, currentValue in arguments:
            if currentArgument in ("-h", "--help"):
                print("usage: main.py [-h]\n")
                print("optional arguments:")
                print("\t-h, --help\tshow this help message")
                print("\t-u, --url <event url>\tuse an ftc events url instead of xlsx files")
                print("\t    --pdf\tsave 6 copies of each sheet in a pdf file") 
                print("\t    --provided-by <name>\tcustomize the provided by field at the top of each table") 
                help = True
            elif currentArgument in ("-u", "--url") and not help:
                print("Accessing {} ...".format(currentValue))
                event_url = currentValue
                api_user = input("Please provide your FTC Events API username: ")
                api_key = input("... and your API key: ")
            # elif currentArgument == "--pdf" and not help:
            #     print ("Enabling pdf mode.")
            #     pdf_mode = True
            elif currentArgument == "--provided-by" and not help:
                provided_by = currentValue
    except getopt.error as err:
        print(str(err))
    else: 
        if not help:
            main(event_url, ":".join((api_user, api_key)), provided_by)