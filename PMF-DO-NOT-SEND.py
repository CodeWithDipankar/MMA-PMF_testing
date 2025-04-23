import itertools
import sys
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import pandas as pd
from datetime import datetime, timedelta
from dateutil.parser import parse
import numpy as np
from pathlib import Path
import os

class LocationDetails:
    def __init__(self, startIndex=None, endIndex=None, noOfWeeks=None):
        self.startIndex = startIndex
        self.endIndex = endIndex
        self.noOfWeeks = noOfWeeks

class ExcelProvider:
    SHEET_NAME: str = "Weekly"

    def excelReader(self, path) -> pd.DataFrame:
        if Path(path).suffix == ".xlsb":
            return pd.read_excel(path, engine='pyxlsb', sheet_name=self.SHEET_NAME, header=8)
        return pd.read_excel(path, sheet_name=self.SHEET_NAME)

    def excelExport(self, path, PMF_FRAMEWORK):
        with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
            for sheetName, _data in PMF_FRAMEWORK.items():
                _data.to_excel(writer, sheet_name=sheetName, index=False)
        print("PMF File Exported!")

    def getWeekLocationForCoreWB(self, columns) -> LocationDetails:
        locs = []
        for loc, colName in enumerate(columns):
            try:
                parse(colName, fuzzy=False)
                locs.append(loc)
            except:
                pass
        if not locs:
            raise ValueError("No week/date columns found in Core Workbook")
        return LocationDetails(min(locs), max(locs) + 1, max(locs) - min(locs) + 1)

    def convertExcelSerialData(self, val):
        try:
            return datetime(1899, 12, 30) + timedelta(days=float(val))
        except:
            return val

    def getWeekLocationForCustomCoreWB(self, columns):
        locs = []
        for loc, col in enumerate(columns):
            val = self.convertExcelSerialData(col)
            if isinstance(val, datetime):
                locs.append(loc)
        if not locs:
            raise ValueError("No week/date columns found in Custom Workbook")
        return LocationDetails(min(locs), max(locs) + 1, max(locs) - min(locs) + 1)

class Controller:
    def run_main_logic(self, core_path, custom_path, no_of_new_weeks, update_ui, on_done):
        try:
            excelProvider = ExcelProvider()
            update_ui("Reading core workbook...")
            CORE_WB_DF = excelProvider.excelReader(core_path)
            CORE_WB_LOC_DETAILS = excelProvider.getWeekLocationForCoreWB(CORE_WB_DF.columns)

            colsRange = list(range(0, 2)) + list(range(CORE_WB_LOC_DETAILS.startIndex, CORE_WB_LOC_DETAILS.endIndex))
            CORE_WEEKLY = CORE_WB_DF.iloc[:, colsRange]
            coreFrameWork = CORE_WEEKLY.iloc[:, :2]

            update_ui("Reading custom workbook...")
            MATCHBACK_C_WB = excelProvider.excelReader(custom_path)
            MATCHBACK_WB_LOC_DETAILS = excelProvider.getWeekLocationForCustomCoreWB(MATCHBACK_C_WB.columns)

            colsRange = list(range(0, 2)) + list(range(MATCHBACK_WB_LOC_DETAILS.startIndex, MATCHBACK_WB_LOC_DETAILS.startIndex + CORE_WB_LOC_DETAILS.noOfWeeks))
            FILTERED_MATCHBACK_C_WB = MATCHBACK_C_WB.iloc[:, colsRange]

            FILTERED_MATCHBACK_C_WB.rename(columns={'Variable Name': 'Variable'}, inplace=True)
            FILTERED_MATCHBACK_C_WB.fillna(0, inplace=True)

            FILTERED_MATCHBACK_C_WB.columns = CORE_WEEKLY.columns
            MATCHBACK_WEEKLY = pd.merge(coreFrameWork, FILTERED_MATCHBACK_C_WB, on=("ModelKey", "Variable"), how='left')

            update_ui("Calculating PMF...")
            PMF = pd.concat([coreFrameWork, (MATCHBACK_WEEKLY.iloc[:, 2:].div(CORE_WEEKLY.iloc[:, 2:], fill_value=1))], axis=1)
            PMF.fillna(1, inplace=True)
            PMF.replace(np.inf, 1, inplace=True)

            for j in range(no_of_new_weeks):
                PMF.iloc[:, len(PMF.columns)-no_of_new_weeks+j] = PMF.iloc[:, len(PMF.columns)-no_of_new_weeks-1]

            PMF[""] = ""
            Actual_Pred = ["Predicted","NNT_GUN","ENT_GUN","SBT_GUN","ROW_GUN","NFB_GUN","NNB_GUN","NNR_GUN","ENB_GUN","NBT_GUN","FSC_DGI","SBC_DGI","STC_DGI","TRC_DGI","SEC_IND"]
            for i in Actual_Pred:
                PMF.loc[PMF['Variable']==i, PMF.columns[2:len(PMF.columns)-1]] = list(itertools.repeat(1, len(PMF.columns)-3))
            PMF["0_Count"] = PMF.isin([0]).sum(axis=1)

            CROSS_CHECK = pd.concat([coreFrameWork,(CORE_WEEKLY.iloc[:,2:].div(MATCHBACK_WEEKLY.iloc[:,2:],fill_value=1))],axis=1)
            CROSS_CHECK.fillna(1,inplace=True)
            CROSS_CHECK[""] = ""
            CROSS_CHECK["0_Count"] = CROSS_CHECK.isin([0]).sum(axis=1)

            ADJ_WEEK = pd.concat([coreFrameWork,(PMF[PMF.columns[2:PMF.shape[1]-2]].multiply(CORE_WEEKLY[CORE_WEEKLY.columns[2:CORE_WEEKLY.shape[1]]]))],axis=1)

            for j in set(ADJ_WEEK['ModelKey']):
                subset = ADJ_WEEK[(ADJ_WEEK['ModelKey'] == j) & ~ADJ_WEEK['Variable'].isin(Actual_Pred)]
                Predicted = subset.sum()
                ADJ_WEEK.loc[(ADJ_WEEK['ModelKey']==j) & (ADJ_WEEK['Variable']=="Predicted"),ADJ_WEEK.columns[2:ADJ_WEEK.shape[1]]] = list(Predicted[2:ADJ_WEEK.shape[1]])

            PMF_FRAMEWORK={
                "Matchback" :MATCHBACK_WEEKLY,
                "Current"   :CORE_WEEKLY,
                "PMF"       :PMF,
                "Cross Check" :CROSS_CHECK,
                "New_Weekly"  : ADJ_WEEK
            }

            update_ui("Saving file...")
            download_dir = str(Path.home() / "Downloads")
            save_path = os.path.join(download_dir, "PMF.xlsx")
            excelProvider.excelExport(save_path, PMF_FRAMEWORK)

            update_ui(f"File saved at:\n{save_path}")

        except Exception as e:
            update_ui(f"Error: {str(e)}")
        finally:
            on_done()

class GUI:
    def __init__(self):
        self.core_path = None
        self.custom_path = None
        self.processing = False
        self.dot_count = 0
        self.ControlManager = Controller()

    def browse_core_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.core_path = path
            self.core_file_display.config(text=Path(path).name)

    def browse_custom_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Binary files", "*.xlsb")])
        if path:
            self.custom_path = path
            self.custom_file_display.config(text=Path(path).name)

    def update_status(self, msg):
        self.status_label.config(text=msg)

    def animate_processing(self):
        if self.processing:
            dots = '.' * (self.dot_count % 4)
            self.status_label.config(text=f"Processing{dots}")
            self.dot_count += 1
            self.root.after(500, self.animate_processing)

    def set_buttons_state(self, state=tk.NORMAL):
        self.browse_core_button.config(state=state)
        self.browse_custom_button.config(state=state)
        self.generate_button.config(state=state)

    def generate(self):
        if not self.core_path or not self.custom_path:
            messagebox.showwarning("Missing file", "Please select both files.")
            return

        try:
            no_of_new_weeks = int(self.week_entry.get())
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter a valid number of new weeks.")
            return

        self.processing = True
        self.dot_count = 0
        self.set_buttons_state(tb.DISABLED)
        self.update_status("Processing")
        self.animate_processing()

        threading.Thread(target=self.ControlManager.run_main_logic, args=(
            self.core_path,
            self.custom_path,
            no_of_new_weeks,
            self.update_status,
            self.process_done
        )).start()

    def process_done(self):
        self.processing = False
        self.set_buttons_state(tb.NORMAL)

    def main(self):
        self.root = tb.Window(themename="cosmo")
        #For ico image
        icon_path = getattr(sys, '_MEIPASS', Path(__file__).parent)
        self.root.iconbitmap(str(Path(icon_path) / "MMA.ico"))
        
        self.root.title("PMF Generator")
        self.root.resizable(False, False)

        frame = tb.Frame(self.root, padding=20)
        frame.pack()

        tb.Label(frame, text="Select Core Workbook (.xlsx):", font=("Helvetica", 11)).grid(row=0, column=0, sticky='w', pady=(0, 10))
        self.core_file_display = tb.Label(frame, text="No file selected", foreground="blue")
        self.core_file_display.grid(row=0, column=1, padx=10, sticky='w')
        self.browse_core_button = tb.Button(frame, text="Browse", command=self.browse_core_file, bootstyle="secondary")
        self.browse_core_button.grid(row=0, column=2)

        tb.Label(frame, text="", font=("Helvetica", 2)).grid(row=1, column=0)

        tb.Label(frame, text="Select Custom Workbook (.xlsb):", font=("Helvetica", 11)).grid(row=2, column=0, sticky='w', pady=(0, 10))
        self.custom_file_display = tb.Label(frame, text="No file selected", foreground="blue")
        self.custom_file_display.grid(row=2, column=1, padx=10, sticky='w')
        self.browse_custom_button = tb.Button(frame, text="Browse", command=self.browse_custom_file, bootstyle="secondary")
        self.browse_custom_button.grid(row=2, column=2)

        # No of New Weeks input
        tb.Label(frame, text="Number of New Weeks:", font=("Helvetica", 11)).grid(row=3, column=0, sticky='w', pady=(10, 10))
        self.week_entry = tb.Entry(frame, width=10)
        self.week_entry.grid(row=3, column=1, sticky='w')

        self.generate_button = tb.Button(frame, text="Generate PMF", command=self.generate, bootstyle="success outline", width=20)
        self.generate_button.grid(row=4, column=0, columnspan=3, pady=25)

        self.status_label = tb.Label(self.root, text="", font=("Helvetica", 10), wraplength=460, justify="left", foreground="#333")
        self.status_label.pack(padx=20, pady=(0, 10), fill="x")

        self.root.mainloop()

if __name__ == "__main__":
    gui = GUI()
    gui.main()
