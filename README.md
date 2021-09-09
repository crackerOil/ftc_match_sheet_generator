# ftc_match_sheet_generator

An individualized match sheet generator for FTC competitions

<h2>How to use</h2>

Requirements: openpyxl, requests

<h3>Using excel files</h3>

1. Put the team names and numbers (preferably taken from [FTC Events](https://ftc-events.firstinspires.org/)) in `teams.xlsx`, in a new folder called `input` (in the same directory as main.py). Example:

![teams.xlsx image](https://i.imgur.com/eHSfCl1.png)

2. Put the matches in `matches.xlsx` in the same folder as `teams.xlsx`. Example:

![matches.xlsx image](https://i.imgur.com/DtbWPyf.png)

3. Run `python main.py`. The match sheets will appear in `output/output.xlsx`

<h3>Using the FTC Events Api</h3>

1. Run `python main.py -u <event url>`

2. Input your API username and key when prompted (sign up for an api token at the [FTC Events API](https://frc-events.firstinspires.org/services/API) page if you don't have one)

3. The match sheets will appear in `output/output.xlsx`
