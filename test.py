from dataclasses import dataclass
from datetime import datetime, date, timedelta
import pandas as pd
import typing as t
from dateutil.parser import parse
path = r"C:\Users\Dipankar.Mandal\OneDrive - Ipsos\MyPersonal\MMA-PMF_testing\Core_Workbook.xlsx"


class LocationDetails:
    def __init__(self, startIndex=None, endIndex=None, noOfWeeks=None):
        self.startIndex = startIndex
        self.endIndex = endIndex
        self.noOfWeeks = noOfWeeks

    def to_dict(self):
        return {
            "startIndex": self.startIndex,
            "endIndex": self.endIndex,
            "noOfWeeks": self.noOfWeeks,
        }

    def __getitem__(self, key):
        return self.to_dict()[key]

    def __repr__(self):
        return str(self.to_dict())

class ExcelProvider:
    locDetails: LocationDetails
    SHEET_NAME: str = "Weekly"

    def excelReader(self, path)->pd.DataFrame:
        if path.lstrip(".") == "xlsb":
            return pd.read_excel(path1, engine='pyxlsb', sheet_name=self.SHEET_NAME, header=8)
        return pd.read_excel(path, sheet_name=self.SHEET_NAME)

    def getWeekLocationForCoreWB(self, columns)->LocationDetails :
        locs = []
        for loc, colName in enumerate(columns):
            try:
                parse(colName, fuzzy=False)
                locs.append(loc)
            except (ValueError, TypeError):
                pass

        if not locs:
            raise ValueError("No week/date columns found in Core Workbook")

        locDetails = LocationDetails(
            startIndex=min(locs),
            endIndex=max(locs),
            noOfWeeks=max(locs) - min(locs) + 1
        )

        return locDetails
    
    def convertExcelSerialData(self, val) -> t.Any:
        try:
            val = float(val)
            return datetime(1899, 12, 30) + timedelta(days=val)
        except:
            return val

    def getWeekLocationForCustomCoreWB(self, columns):
        locs = []
        for loc, col in enumerate(columns):
            if isinstance(col, (int, float)):
                converted = self.convertExcelSerialData(col)
                if isinstance(converted, datetime):
                    locs.append(loc)
                    continue
            if isinstance(col, (datetime, date)):
                locs.append(loc)
        
        if not locs:
            raise ValueError("No week/date columns found in Custom Workbook")

        locDetails = LocationDetails(
            startIndex=min(locs),
            endIndex=max(locs),
            noOfWeeks=max(locs) - min(locs) + 1
        )

        return locDetails


class PMFInitiator:
    CUSTOM_WB_PATH: str = ""
    CORE_WB_PATH: str = ""


if __name__ == "__main__":
    excelProvider = ExcelProvider()
    CORE_WB_DF = excelProvider.excelReader(path)
    CORE_WB_LOC_DETAILS = excelProvider.getWeekLocationForCoreWB(CORE_WB_DF.columns)

    colsRange = list(range(0, 2)) + list(range(CORE_WB_LOC_DETAILS.startIndex, CORE_WB_LOC_DETAILS.noOfWeeks))
    CORE_NEW_WB_DF = CORE_WB_DF.iloc[:,colsRange]
    coreFrameWork = CORE_NEW_WB_DF.iloc[:,:2]
    del CORE_WB_DF

    MATCHBACK_C_WB = excelProvider.excelReader(path1)
    MATCHBACK_WB_LOC_DETAILS = excelProvider.getWeekLocationForCustomCoreWB(MATCHBACK_C_WB.columns)
    
    colsRange = list(range(0, 2)) + list(range(MATCHBACK_WB_LOC_DETAILS.startIndex, CORE_WB_LOC_DETAILS.noOfWeeks))
    FILTERED_MATCHBACK_C_WB = MATCHBACK_C_WB.iloc[:,colsRange]

    FILTERED_MATCHBACK_C_WB.rename(columns={'Variable Name':'Variable'},inplace=True)
    FILTERED_MATCHBACK_C_WB.fillna(0,inplace=True)

    # ## MatchBack TAb and Current Tab ##
    # NEW_WEEKLY.columns=Matchback_CWB.columns
    # NEW_WEEKLY
    # Matchback_Weekly=pd.merge(Framework,Matchback_CWB,on=("ModelKey","Variable"),how='left')
    # del Matchback_CWB
    # Matchback_Weekly
