# ftc_match_sheet_generator

An individualized match sheet generator for FTC competitions

<h2>How to use</h2>

<h3>Using excel files</h3>

1. Install openpyxl with `pip install openpyxl`

2. Put the team names and numbers (preferably taken from [FTC Events](https://ftc-events.firstinspires.org/)) in `teams.xlsx`, in a new folder called `excel` (in the same directory as main.py). Example:

![teams.xlsx image](https://i.imgur.com/eHSfCl1.png)

3. Put the matches in `matches.xlsx` in the same folder as `teams.xlsx`. Example:

![matches.xlsx image](https://i.imgur.com/DtbWPyf.png)

4. Run `python main.py`. The match sheets will appear in `excel/output.xlsx`

<h3>Using the FTC Events Api</h3>

1. Install openpyxl with `pip install openpyxl`

2. Run `python main.py -u <event url>`

3. Input your API username and key when prompted (sign up for an api token at the [FTC Events API](https://frc-events.firstinspires.org/services/API) page if you don't have one)

4. The match sheets will appear in `excel/output.xlsx`
