from dataclasses import dataclass
from datetime import datetime, date, timedelta
import itertools
import os
from pathlib import Path
import numpy as np
import pandas as pd
import typing as t
from dateutil.parser import parse

path = r"C:\Users\Dipankar.Mandal\OneDrive - Ipsos\MyPersonal\MMA-PMF_testing\Core_Workbook.xlsx"
path1 = r"C:\Users\Dipankar.Mandal\OneDrive - Ipsos\MyPersonal\MMA-PMF_testing\CustomWorkbook.xlsb"
exportPath = r"C:\Users\Dipankar.Mandal\OneDrive - Ipsos\MyPersonal\MMA-PMF_testing\PMF.xlsx"
n = 4

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
        if Path(path).suffix == ".xlsb":
            return pd.read_excel(path1, engine='pyxlsb', sheet_name=self.SHEET_NAME, header=8)
        return pd.read_excel(path, sheet_name=self.SHEET_NAME)
    
    def excelExport(self, path, PMF_FRAMEWORK):
        with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
            for sheetName, _data in PMF_FRAMEWORK.items():
                _data.to_excel(writer, sheet_name=sheetName, index=False)

        print("PMF File Exported!")

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
            endIndex=max(locs) + 1,
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
            endIndex=max(locs) + 1,
            noOfWeeks=max(locs) - min(locs) + 1
        )

        return locDetails


if __name__ == "__main__":
    excelProvider = ExcelProvider()
    CORE_WB_DF = excelProvider.excelReader(path)
    CORE_WB_LOC_DETAILS = excelProvider.getWeekLocationForCoreWB(CORE_WB_DF.columns)

    colsRange = list(range(0, 2)) + list(range(CORE_WB_LOC_DETAILS.startIndex, CORE_WB_LOC_DETAILS.endIndex))
    CORE_WEEKLY = CORE_WB_DF.iloc[:,colsRange]
    coreFrameWork = CORE_WEEKLY.iloc[:,:2]
    del CORE_WB_DF

    MATCHBACK_C_WB = excelProvider.excelReader(path1)
    MATCHBACK_WB_LOC_DETAILS = excelProvider.getWeekLocationForCustomCoreWB(MATCHBACK_C_WB.columns)
    
    colsRange = list(range(0, 2)) + list(range(MATCHBACK_WB_LOC_DETAILS.startIndex, MATCHBACK_WB_LOC_DETAILS.startIndex + CORE_WB_LOC_DETAILS.noOfWeeks))
    FILTERED_MATCHBACK_C_WB = MATCHBACK_C_WB.iloc[:,colsRange]

    FILTERED_MATCHBACK_C_WB.rename(columns={'Variable Name':'Variable'},inplace=True)
    FILTERED_MATCHBACK_C_WB.fillna(0,inplace=True)

    FILTERED_MATCHBACK_C_WB.columns = CORE_WEEKLY.columns
    MATCHBACK_WEEKLY = pd.merge(coreFrameWork,FILTERED_MATCHBACK_C_WB,on=("ModelKey","Variable"),how='left')
    del MATCHBACK_C_WB
    del FILTERED_MATCHBACK_C_WB

    PMF = pd.concat([coreFrameWork,(MATCHBACK_WEEKLY.iloc[:,2:].div(CORE_WEEKLY.iloc[:,2:],fill_value=1))],axis=1)
    PMF.fillna(1,inplace=True)
    PMF.replace(np.inf,1,inplace=True)

    for j in range(n):
        PMF.iloc[:,len(PMF.columns)-n+j]=PMF.iloc[:,len(PMF.columns)-n-1]
    
    PMF[""] = ""
    Actual_Pred = ["Predicted","NNT_GUN","ENT_GUN","SBT_GUN","ROW_GUN","NFB_GUN","NNB_GUN","NNR_GUN","ENB_GUN","NBT_GUN","FSC_DGI","SBC_DGI","STC_DGI","TRC_DGI","SEC_IND"]
    for i in Actual_Pred:
        PMF.loc[PMF['Variable']==i,PMF.columns[2:len(PMF.columns)-1]]= list(itertools.repeat(1,len(PMF.columns)-3))
    PMF["0_Count"]=PMF.isin([0]).sum(axis=1)

    ## Cross Check TAB ##
    CROSS_CHECK = pd.concat([coreFrameWork,(CORE_WEEKLY.iloc[:,2:].div(MATCHBACK_WEEKLY.iloc[:,2:],fill_value=1))],axis=1)
    CROSS_CHECK.fillna(1,inplace=True)
    CROSS_CHECK[""] = ""
    CROSS_CHECK["0_Count"]=CROSS_CHECK.isin([0]).sum(axis=1)

    ## NEW_WEEKLY TAB ##
    ADJ_WEEK = pd.concat([coreFrameWork,(PMF[PMF.columns[2:PMF.shape[1]-2]].multiply(CORE_WEEKLY[CORE_WEEKLY.columns[2:CORE_WEEKLY.shape[1]]]))],axis=1)

    for j in set(ADJ_WEEK['ModelKey']):
        subset = ADJ_WEEK[(ADJ_WEEK['ModelKey'] == j) & ~ADJ_WEEK['Variable'].isin(["Predicted","NNT_GUN","ENT_GUN","SBT_GUN","ROW_GUN","NFB_GUN","NNB_GUN","NNR_GUN","ENB_GUN","NBT_GUN"])]
        Predicted = subset.sum()
        ADJ_WEEK.loc[(ADJ_WEEK['ModelKey']==j) & (ADJ_WEEK['Variable']=="Predicted"),ADJ_WEEK.columns[2:ADJ_WEEK.shape[1]]]=list(Predicted[2:ADJ_WEEK.shape[1]])

    ## EXPORT  ##
    PMF_FRAMEWORK={
        "Matchback" :MATCHBACK_WEEKLY,
        "Current"   :CORE_WEEKLY,
        "PMF"       :PMF,
        "Cross Check" :CROSS_CHECK,
        "New_Weekly"  : ADJ_WEEK
    }
    
    excelProvider.excelExport(exportPath, PMF_FRAMEWORK)