
import pandas as pd

from dateutil.parser import parse
path = r"C:\Users\Dipankar.Mandal\OneDrive - Ipsos\MyPersonal\MMA-PMF_testing\Core_Workbook.xlsx"


def weekLocation(columns):
    locs = []
    for loc , colNmae in enumerate(columns):
        try:
            parse(colNmae, fuzzy=False)
            locs.append(loc) 
        except (ValueError, TypeError):
            pass
    return locs

if __name__ == "__main__":
    df = pd.read_excel(path, sheet_name="Weekly")
    columns = df.columns
    locs = weekLocation(columns)
    startIndex, noOfWeeks = min(locs), (max(locs)-min(locs)+1)

    print()