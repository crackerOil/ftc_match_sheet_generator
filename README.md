# ftc_match_sheet_generator

An individualized match sheet generator for FTC competitions

<h2>How to use</h2>

1. Install openpyxl with `pip install openpyxl`
2. Put the team names and numbers (preferably taken from [FTC Events](https://ftc-events.firstinspires.org/)) in `teams.xlsx`, in a new folder called `excel` (in the same directory as main.py). Example:

![teams.xlsx image](https://i.imgur.com/eHSfCl1.png)

3. Put the matches in `matches.xlsx` in the same folder as `teams.xlsx`. Example:

![matches.xlsx image](https://i.imgur.com/DtbWPyf.png)

4. Run `python main.py`. The match sheets will appear in `output.xlsx`
