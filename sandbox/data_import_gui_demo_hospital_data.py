from datetime import date, datetime, timedelta
import glob
import os
import tkinter as tk
from tkinter import filedialog
import time

import fpdf
import pandas as pd
from pandas_schema import Column, Schema
from pandas_schema.validation import InListValidation
import xlsxwriter
import tabulate

import create_PDF_report as report

class Application(tk.Frame):
    def __init__(self, titled, greeting, master=None):
        super().__init__(master)
        self.titled = titled
        self.greeting = greeting
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        """Create initial window and initiate the upload process"""
        self.master.title(self.titled)
        self.hi_there = tk.Label(self)
        self.hi_there["text"] = self.greeting
        self.hi_there.pack(side="top")
        self.next = tk.Button(self, text="CHOOSE FOLDER", fg="green", bg='#acacad',
                              command=self.load_data)
        self.next.pack(side="right")
        self.quit = tk.Button(self, text="QUIT", fg="red",
                              command=self.master.destroy)
        self.quit.pack(side="left")

    def load_data(self):
        """Open a dialog box to select the file directory that contains data to be imported"""
        self.dirname = filedialog.askdirectory()
        search_dir = self.dirname + "/" + "*Hospital Capacity*"
        self.master.destroy()
        self.new_window = tk.Tk()
        self.new_window.title(self.titled)
        self.file_list = glob.glob(search_dir)
        all_files = ""
        for file in self.file_list:
            all_files = all_files + " " + file + "\n"
        blurb = tk.Label(self.new_window)
        blurb["text"] = "Here are the files you selected for cleaning + compiling: \n\n "\
                        + all_files + "\n Note: please make sure none of these files are open \n\n "\
                        "If you are appending the new data to a historical dataset, click LOAD HISTORICAL DATASET. \n\n \
                        If there is no historical data, click RUN DATA IMPORT"
        blurb.pack(side="top")
        next = tk.Button(self.new_window, text="LOAD HISTORICAL DATASET", fg="green", bg='#acacad',
                         command=self.load_historical)
        next.pack(side="right")
        self.quit = tk.Button(self.new_window, text="QUIT", fg="red", bg='#acacad',
                              command=self.new_window.destroy)
        self.quit.pack(side="left")
        jump_to_clean = tk.Button(self.new_window, text="RUN DATA IMPORT", fg="blue", bg='#acacad',
                                  command=self.clean_data)
        jump_to_clean.pack(side="left")


    def load_historical(self):
        """Option step where you include previously cleaned historical data"""
        self.historical_file = filedialog.askopenfilename()
        self.new_window.destroy()
        self.load_next = tk.Tk()
        self.load_next.title(self.titled)
        self.text = tk.Label(self.load_next)
        self.text["text"] = "Here is the historical file you selected:\n" + self.historical_file + \
                            " \n choose RUN DATA IMPORT to continue the data import process"
        self.text.pack(side="top")
        self.next = tk.Button(self.load_next, text="RUN DATA IMPORT", fg="green", bg='#acacad',
                              command=self.clean_data)
        self.next.pack(side="right")
        self.quit = tk.Button(self.load_next, text="QUIT", fg="red", bg='#acacad',
                              command=self.load_next.destroy)
        self.quit.pack(side="left")

    def clean_data(self):
        """initialize the data frame structure with the first file and cycle through
        the remaining files in the file list and append them to the "starter" file"""
        try:
            self.load_next.destroy()
        except AttributeError:
            self.new_window.destroy()
        self.hospital_names = ['Abington Memorial Hospital', 'Albert Einstein at Elkins Park',
                             'Einstein Medical Center Hospital',
                             'Holy Redeemer Hospital & Medical Center',
                             'Main Line-Bryn Mawr Hospital',
                             'Main Line-Lankenau Hospital', 'Suburban Community Hospital-Norristown',
                             'Pottstown Memorial Medical Center', 'Lansdale Hospital']

        # TODO: This is throwing an error sometimes, but not others
        # 'IndexError: pop from empty list'
        initial_file = self.file_list.pop()
        df = pd.read_excel(initial_file, skiprows=3)
        df['source_file'] = initial_file.split("/")[-1]
        for file in self.file_list:
            df_next = pd.read_excel(file, skiprows=3)
            df_next['source_file'] = file.split("/")[-1]
            df = df.append(df_next)

        # Clean up and rename columns/rows
        self.data = df[df['Unnamed: 0'] != 'Total'].copy()
        # Drop the first column b/c it's empty
        self.data.drop('Unnamed: 1', axis=1, inplace=True)
        # Section 1.A: Change columns here. Column order, zero-indexed (self.data.columns[0] is the first column)
        # column names follow SQL Server naming conventions (50 character limit, CamelCase, no symbols, limit abbreviation)
        self.data.rename(columns={self.data.columns[0]: "HospitalName",
                                  self.data.columns[1]: "ICUStaffedBed",
                                  self.data.columns[2]: "ICUAvailableNow",
                                  self.data.columns[3]: "ICUAvailable24H",
                                  self.data.columns[4]: "ICUAvailable72H",
                                  self.data.columns[5]: "MedSurgStaffedBed",
                                  self.data.columns[6]: "MedSurgAvailableNow",
                                  self.data.columns[7]: "MedSurgAvailable24H",
                                  self.data.columns[8]: "MedSurgAvailable72H",
                                  self.data.columns[9]: "BurnStaffedBed",
                                  self.data.columns[10]: "BurnAvailableNow",
                                  self.data.columns[11]: "BurnAvailable24H",
                                  self.data.columns[12]: "BurnAvailable72H",
                                  self.data.columns[13]: "PedsICUStaffedBed",
                                  self.data.columns[14]: "PedsICUAvailableNow",
                                  self.data.columns[15]: "PedsICUAvailable24H",
                                  self.data.columns[16]: "PedsICUAvailable72H",
                                  self.data.columns[17]: "PedsStaffedBed",
                                  self.data.columns[18]: "PedsAvailableNow",
                                  self.data.columns[19]: "PedsAvailable24H",
                                  self.data.columns[20]: "PedsAvailable72H",
                                  self.data.columns[21]: "NeonatalStaffedBed",
                                  self.data.columns[22]: "NeonatalAvailableNow",
                                  self.data.columns[23]: "NeonatalAvailable24H",
                                  self.data.columns[24]: "NeonatalAvailable72H",
                                  self.data.columns[25]: "InpatientRehabStaffedBed",
                                  self.data.columns[26]: "InpatientRehabAvailableNow",
                                  self.data.columns[27]: "InpatientRehabAvailable24H",
                                  self.data.columns[28]: "InpatientRehabAvailable72H",
                                  self.data.columns[29]: "PyschStaffedBed",
                                  self.data.columns[30]: "PyschAvailableNow",
                                  self.data.columns[31]: "PyschAvailable24H",
                                  self.data.columns[32]: "PyschAvailable72H",
                                  self.data.columns[33]: "PyschAdultStaffedBed",
                                  self.data.columns[34]: "PyschAdultAvailableNow",
                                  self.data.columns[35]: "PyschAdultAvailable24H",
                                  self.data.columns[36]: "PyschAdultAvailable72H",
                                  self.data.columns[37]: "PyschAdolStaffedBed",
                                  self.data.columns[38]: "PyschAdolAvailableNow",
                                  self.data.columns[39]: "PyschAdolAvailable24H",
                                  self.data.columns[40]: "PyschAdolAvailable72H",
                                  self.data.columns[41]: "PyschGeriStaffedBed",
                                  self.data.columns[42]: "PyschGeriAvailableNow",
                                  self.data.columns[43]: "PyschGeriAvailable24H",
                                  self.data.columns[44]: "PyschGeriAvailable72H",
                                  self.data.columns[45]: "PyschDetoxStaffedBed",
                                  self.data.columns[46]: "PyschDetoxAvailableNow",
                                  self.data.columns[47]: "PyschDetoxAvailable24H",
                                  self.data.columns[48]: "PyschDetoxAvailable72H",
                                  self.data.columns[49]: "PyschSustanceDualStaffedBed",
                                  self.data.columns[50]: "PyschSustanceDualAvailableNow",
                                  self.data.columns[51]: "PyschSustanceDualAvailable24H",
                                  self.data.columns[52]: "PyschSustanceDualAvailable72H",
                                  self.data.columns[53]: "LaborDeliverStaffedBed",
                                  self.data.columns[54]: "LaborDeliverAvailableNow",
                                  self.data.columns[55]: "LaborDeliverAvailable24H",
                                  self.data.columns[56]: "LaborDeliverAvailable72H",
                                  self.data.columns[57]: "MaternityStaffedBed",
                                  self.data.columns[58]: "MaternityAvailableNow",
                                  self.data.columns[59]: "MaternityAvailable24H",
                                  self.data.columns[60]: "MaternityAvailable72H",
                                  self.data.columns[61]: "AirborneIsoStaffedBed",
                                  self.data.columns[62]: "AirborneIsoAvailableNow",
                                  self.data.columns[63]: "AirborneIsoAvailable24H",
                                  self.data.columns[64]: "AirborneIsoAvailable 72H",
                                  self.data.columns[65]: "EDImmediate",
                                  self.data.columns[66]: "EDDelayed",
                                  self.data.columns[67]: "EDMinor",
                                  self.data.columns[68]: "EDDeceased",
                                  self.data.columns[69]: "NumPatientWaitingNonCOVIDAdmit",
                                  self.data.columns[70]: "NumPatientNonVentCOVIDAdmit",
                                  self.data.columns[71]: "NumPatientVentCOVIDAdmit",
                                  self.data.columns[72]: "NumPatientWaitingICUBed",
                                  self.data.columns[73]: "NumPatientWaitingDischarge",
                                  self.data.columns[74]: "YesterdayCOVIDAdmit",
                                  self.data.columns[75]: "YesterdayPUIAdmit",
                                  self.data.columns[76]: "RespiratoryProtectionPlanIndicator",
                                  self.data.columns[77]: "N95PlanFitTested",
                                  self.data.columns[78]: "N95BrandModel",
                                  self.data.columns[79]: "PARPsPlanTrained",
                                  self.data.columns[80]: "PPETrainStatus",
                                  self.data.columns[81]: "NeedHandSanitizer",
                                  self.data.columns[82]: "NeedHandSoap",
                                  self.data.columns[83]: "NeedDisinfectionSolution",
                                  self.data.columns[84]: "NeedDisinfectionWipes",
                                  self.data.columns[85]: "NeedGloves",
                                  self.data.columns[86]: "NeedOther",
                                  self.data.columns[87]: "ExpectShortageN95",
                                  self.data.columns[88]: "ExpectShortagePARP",
                                  self.data.columns[89]: "ExpectShortagePARPHood",
                                  self.data.columns[90]: "ExpectShortagePARPFilter",
                                  self.data.columns[91]: "ExpectShortageFacialMask",
                                  self.data.columns[92]: "ExpectShortageGownApron",
                                  self.data.columns[93]: "ExpectShortageEyeProtection",
                                  self.data.columns[94]: "ExpectShortageDisinfectionSupply",
                                  self.data.columns[95]: "ExpectShortageOther",
                                  self.data.columns[96]: "COVIDResExpectShortageN95",
                                  self.data.columns[97]: "COVIDResExpectShortagePARP",
                                  self.data.columns[98]: "COVIDResExpectShortagePARPHood",
                                  self.data.columns[99]: "COVIDResExpectShortagePARPFilter",
                                  self.data.columns[100]: "COVIDResExpectShortageFacialMask",
                                  self.data.columns[101]: "COVIDResExpectShortageGown",
                                  self.data.columns[102]: "COVIDResExpectShortageEyeProtection",
                                  self.data.columns[103]: "COVIDResExpectShortageHandSoap",
                                  self.data.columns[104]: "COVIDResExpectShortageHandSanitizer",
                                  self.data.columns[105]: "COVIDResExpectShortageDisinfectionSupply",
                                  self.data.columns[106]: "COVIDResExpectShortageOther",
                                  self.data.columns[107]: "N95BurnRate",
                                  self.data.columns[108]: "PARPBurnRate",
                                  self.data.columns[109]: "PARPHoodBurnRate",
                                  self.data.columns[110]: "PARPFilterBurnRate",
                                  self.data.columns[111]: "FacialMaskBurnRate",
                                  self.data.columns[112]: "GownBurnRate",
                                  self.data.columns[113]: "EyeProtectionBurnRate",
                                  self.data.columns[114]: "AnticipatedShortageTestingCollection",
                                  self.data.columns[115]: "ShortageNote",
                                  self.data.columns[116]: "IndicateCommercialOrInhouseCOVIDTesting",
                                  self.data.columns[117]: "TestingGoLiveDate",
                                  self.data.columns[118]: "COVIDTestRunInhouseToday",
                                  self.data.columns[119]: "COVIDPositiveTestToday",
                                  self.data.columns[120]: "TotalInpatientCOVIDDiagnosed",
                                  self.data.columns[121]: "TotalInpatientPUI",
                                  self.data.columns[122]: "TotalICUBedOccupiedByCOVIDDiagnosed",
                                  self.data.columns[123]: "YesterdayInpatientAdmitOver14DayConvertCOVID",
                                  self.data.columns[124]: "TotalInpatientAdmitOver14DayConvertCOVID",
                                  self.data.columns[125]: "TotalCOVIDDiagnosedOnVent",
                                  self.data.columns[126]: "TotalCOVIDDiagnosedOnECMO",
                                  self.data.columns[127]: "NumAirborneInfectionIsoED",
                                  self.data.columns[128]: "NumAirborneInfectionIsoICU",
                                  self.data.columns[129]: "NumAirborneInfectionIsoNonICU",
                                  self.data.columns[130]: "24HCOVIDPatientDeath",
                                  self.data.columns[131]: "Yesterday24HCOVIDPatientDeath",
                                  self.data.columns[132]: "IndicateExtendedRespiratorUse",
                                  self.data.columns[133]: "IndicateReusableRespiratorUse",
                                  self.data.columns[134]: "IndicateReuseN95",
                                  self.data.columns[135]: "IndicateExtendedStaffHours",
                                  self.data.columns[136]: "IndicateNumCohortingNoDedicatedStaff",
                                  self.data.columns[137]: "IndicateNumCohortingDedicatedStaff",
                                  self.data.columns[138]: "IndicateN95Last1to3Day",
                                  self.data.columns[139]: "IndicateN95Last4to14Day",
                                  self.data.columns[140]: "IndicateN95LastOver14Day",
                                  self.data.columns[141]: "IndicatePPELast1to3Day",
                                  self.data.columns[142]: "IndicatePPELast4to14Day",
                                  self.data.columns[143]: "IndicatePPELastOver14Day",
                                  self.data.columns[144]: "IndicateN95SpecimenLast1to3Day",
                                  self.data.columns[145]: "IndicateN95SpecimenLast4to14Day",
                                  self.data.columns[146]: "IndicateN95SpecimenLastOver14Day",
                                  self.data.columns[147]: "TotalEmployeeAbsent",
                                  self.data.columns[148]: "EmployeeAbsentCOVID",
                                  self.data.columns[149]: "PhysicianCallOut",
                                  self.data.columns[150]: "NurseCallOut",
                                  self.data.columns[151]: "ExposureCallOut",
                                  self.data.columns[152]: "ChildCareCallOut",
                                  self.data.columns[153]: "CriticalStaffShortageEnvScience",
                                  self.data.columns[154]: "CriticalStaffShortageRNandLPN",
                                  self.data.columns[155]: "CriticalStaffShortageRespTherapist",
                                  self.data.columns[156]: "CriticalStaffShortagePharma",
                                  self.data.columns[157]: "CriticalStaffShortagePhysician",
                                  self.data.columns[158]: "CriticalStaffShortageOtherLicensedIP",
                                  self.data.columns[159]: "CriticalStaffShortageTemporaryPhysicianAndLP",
                                  self.data.columns[160]: "CriticalStaffShortageOtherHCP",
                                  self.data.columns[161]: "CriticalStaffShortageNotListed",
                                  self.data.columns[162]: "InAWeekCriticalStaffShortageEnvScience",
                                  self.data.columns[163]: "InAWeekCriticalStaffShortageRNandLPN",
                                  self.data.columns[164]: "InAWeekCriticalStaffShortageRespTherapist",
                                  self.data.columns[165]: "InAWeekCriticalStaffShortagePharma",
                                  self.data.columns[166]: "InAWeekCriticalStaffShortagePhysician",
                                  self.data.columns[167]: "InAWeekCriticalStaffShortageOtherLicensedIP",
                                  self.data.columns[168]: "InAWeekCriticalStaffShortageTemporaryPhysicianAndLP",
                                  self.data.columns[169]: "InAWeekCriticalStaffShortageOtherHCP",
                                  self.data.columns[170]: "InAWeekCriticalStaffShortageNotListed",
                                  self.data.columns[171]: "NumVentilator",
                                  self.data.columns[172]: "NumVentilatorsInUse",
                                  self.data.columns[173]: "NumAnesthesiaMachines",
                                  self.data.columns[174]: "NumAnesthesiaMachinesConvertedToVent",
                                  self.data.columns[175]: "NumVentsUsedForConfirmedCOVIDPatient",
                                  self.data.columns[176]: "NumECMOUnit",
                                  self.data.columns[177]: "NumECMOInUse",
                                  self.data.columns[178]: "NumECMOInUseForCOVIDPatient",
                                  self.data.columns[179]: "TotalEDAirborneIsolationRoom",
                                  self.data.columns[180]: "AvailableEDAirborneIsolationRoom",
                                  self.data.columns[181]: "EDAirborneIsolationOccupiedReqIsolation",
                                  self.data.columns[182]: "EDAirborneIsolationOccupiedByCOVID",
                                  self.data.columns[183]: "TotalNonICUAirborneIsolationRoom",
                                  self.data.columns[184]: "AvailableNonICUAirborneIsolationRoom",
                                  self.data.columns[185]: "NonICUAirborneIsolationOccupiedReqIsolation",
                                  self.data.columns[186]: "NonICUAirborneIsolationOccupiedByCOVID",
                                  self.data.columns[187]: "TotalICUAirborneIsolationRoom",
                                  self.data.columns[188]: "AvailableICUAirborneIsolationRoom",
                                  self.data.columns[189]: "ICUAirborneIsolationOccupiedReqIsolation",
                                  self.data.columns[190]: "ICUAirborneIsolationOccupiedByCOVID",
                                  self.data.columns[191]: "EDDivertStatus",
                                  self.data.columns[192]: "MassDeconStatus",
                                  self.data.columns[193]: "VentFullFeature",
                                  self.data.columns[194]: "VentPediatricCapable",
                                  self.data.columns[195]: "VentRescueTherapy",
                                  self.data.columns[196]: "FacilityStress1",
                                  self.data.columns[197]: "FacilityStress2",
                                  self.data.columns[198]: "FacilityStress3",
                                  self.data.columns[199]: "FacilityStress4",
                                  self.data.columns[200]: "FacilityStress5",
                                  self.data.columns[201]: "EmergencyGeneratorAvailable",
                                  self.data.columns[202]: "HeatingAndACUnderGeneratorAvailable",
                                  self.data.columns[203]: "PhoneAvailableUnderGenerator",
                                  self.data.columns[204]: "HeatingSystemElectric",
                                  self.data.columns[205]: "HeatingSystemNaturalGas",
                                  self.data.columns[206]: "HeatingSystemOil",
                                  self.data.columns[207]: "HeatingSystemPropane",
                                  self.data.columns[208]: "TotalAvailableNonSkilledAmbulatoryBed",
                                  self.data.columns[209]: "TotalAvailableNonSkilledAmbulatory",
                                  self.data.columns[210]: "TotalAvailableNonSkilledNonAmbulatoryBed",
                                  self.data.columns[211]: "TotalAvailableNonSkilledNonAmbulatory",
                                  self.data.columns[212]: "TotalAvailableSkilledBed",
                                  self.data.columns[213]: "TotalAvailableSkilled",
                                  self.data.columns[214]: "SkilledFeedingTubeAvailable",
                                  self.data.columns[215]: "SkilledIVFluidAvailable",
                                  self.data.columns[216]: "SkilledIsolationAvailable",
                                  self.data.columns[217]: "SkilledSecurityAvailable",
                                  self.data.columns[218]: "SkilledVentilatorAvailable",
                                  self.data.columns[219]: "TotalFTCertifiedNursingAssistant1",
                                  self.data.columns[220]: "TotalPTCertifiedNursingAssistant1",
                                  self.data.columns[221]: "TotalFTCertifiedNursingAssistant2",
                                  self.data.columns[222]: "TotalPTCertifiedNursingAssistant2",
                                  self.data.columns[223]: "TotalFTLPN",
                                  self.data.columns[224]: "TotalPTLPN",
                                  self.data.columns[225]: "TotalFTMedicalTechnician",
                                  self.data.columns[226]: "TotalPTMedicalTechnician",
                                  self.data.columns[227]: "TotalFTPharmacist",
                                  self.data.columns[228]: "TotalPTPharmacist",
                                  self.data.columns[229]: "TotalFTRN",
                                  self.data.columns[230]: "TotalPTRN",
                                  self.data.columns[231]: "TotalFTSocialService",
                                  self.data.columns[232]: "TotalPTSocialService",
                                  self.data.columns[233]: "TotalOnsiteFeedingTubePump",
                                  self.data.columns[234]: "TotalOnsiteHospitalBed",
                                  self.data.columns[235]: "TotalOnsiteIVPump",
                                  self.data.columns[236]: "TotalOnsiteStationaryBed",
                                  self.data.columns[237]: "TotalOnsiteVentilator"
                                  },
                    inplace=True)

        # Add some new columns
        self.data["date"] = self.data['source_file'].replace(regex={'Hospital Capacity and EEI ': '', '.xlsx': ''})
        self.data["UploadCount"] = 1
        self.data["ImportStatus"] = 'current'
        cols = ['HospitalName', "date", "UploadCount"]
        self.data["UniqueID"] = self.data[cols].apply(lambda row: '_'.join(row.values.astype(str)), axis=1)

        # subset data for key hospitals
        self.hospitals = self.data[self.data['HospitalName'].isin(self.hospital_names)]
        try:
            historical_data = pd.read_csv(self.historical_file)
            historical_data['Source2'] = "previous"
            self.hospitals = self.hospitals.append(historical_data)
        except AttributeError:
            pass

        # Write out clean data
        save_path = self.dirname + '/cleaned_data/'
        if not os.path.exists(save_path):
            os.makedirs(save_path)
        self.data.to_csv(path_or_buf= save_path + "Compiled_All_Hospital_Data.csv", index=False, na_rep="NULL")
        self.hospitals.to_csv(path_or_buf= save_path + "Compiled_County_Hospital_Data.csv", index=False, na_rep="NULL")

        self.validate_and_annotate(save_path + "Compiled_All_Hospital_Data.csv")
        

    def validate_and_annotate(self, file_path=''):
        ################################
        #   SCHEMAS VALIDATION    #
        ################################

        # FIX ME: this is a really janky way to do this. Passing empty string for file_path...
        # is a temp fix for review_changes re-validate data button causing an infinite loop when calling this with an argument :(
        if not file_path:
            file_path = self.dirname + '/for_review/intermediate_hospital_data.xlsx'
        
        try:
            df = pd.read_csv(file_path)
        except:
            try:
                df = pd.read_excel(file_path)
            except:
                print("UNACCEPTED FILE FORMAT")

        df.fillna("NULL", inplace= True)
        
        #~ print(file_path)
        #~ print(df.info)

        schema = Schema([
            Column('N95PlanFitTested', [InListValidation(['Y', 'N', 'NULL'])]),
            Column('PARPsPlanTrained', [InListValidation(['Y', 'N', 'NULL'])]),
            ])

        errors = schema.validate(df, columns=schema.get_column_names())
            
        #######################################
        # Build excel worksheet w formatting
        #######################################
        save_path2 = self.dirname + '/for_review/'
        if not os.path.exists(save_path2):
            os.makedirs(save_path2)

        writer = pd.ExcelWriter(save_path2 +'intermediate_hospital_data.xlsx', engine='xlsxwriter')

        # Skip row 1 headers so we can add manunally with formatting
        df.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False, index=False)

        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        ### WORKBOOK FORMATS ###
        yellow_highlight = workbook.add_format({'bg_color': '#FFEB9C' })

        header = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
            })
        ########################

        # Set column widths
        worksheet.set_column('A:II', 30)
        worksheet.set_default_row(hide_unused_rows=True)

        # Write the column headers with the defined format.
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header)

        # for storing error row numbers while we iterate thru error object
        # will use for hiding rows 
        error_rows = []
        df_length = len(df)
        
        for error in errors:

            error_rows.append(error.row)
            
            row = error.row + 1
            column = df.columns.get_loc(error.column)
            
            # Comments
            worksheet.write_comment(row, column , error.message)

            # Highlights
            worksheet.conditional_format(row, column, row, column, {'type': 'no_errors', 'format': yellow_highlight})

        #~ print(error_rows);

        # Hide Rows that don't contain errors
        for i in range(df_length+1):
            if i not in error_rows:
                worksheet.set_row(i + 1, None, None, {'hidden': True})
        

        writer.save()

        # Pop up
        self.review_changes()


    ## Replace compiled data file with edited intermediate file
    def use_intermediate_file(self):
        read_path = self.dirname + '/for_review/intermediate_hospital_data.xlsx'
        df = pd.read_excel(read_path, skip_blank_lines=True)
        df.dropna
        #~ print(df.info)

        save_path = self.dirname + '/cleaned_data/'
        if not os.path.exists(save_path):
            os.makedirs(save_path)
        df.to_csv(path_or_buf= save_path + "Compiled_All_Hospital_Data.csv", index=False, na_rep="NULL")

        ## Compiled_All_Hospital_Data.csv

        self.cleaned_message()
        


    def review_changes(self):
        modified_file = self.dirname + '/for_review/intermediate_hospital_data.xlsx'
        if hasattr(self, 'review_window'):
            self.review_window.destroy()
        self.review_window = tk.Tk()
        self.review_window.title(self.titled)
        note = tk.Label(self.review_window)
        note["text"] = "Some data entries have been modified or contain suggested changes. \n" \
                        "You can review these in /test_data/intermediate_hospital_data.xlsx \n \n" \
                        "Once you are finished, save the intermediate file and continue"
        note.pack(side="top")

        ## Use unchanged file
        use_original = tk.Button(self.review_window, text="Ignore Changes", fg="red", bg='#acacad',
                        command = self.cleaned_message)
        use_original.pack(side="left")

        ## Re-Validate file
        re_validate = tk.Button(self.review_window, text="Re-Check Data", fg="black", bg='#acacad',
                        command = self.validate_and_annotate)
        re_validate.pack(side="left")

        ## Use intermediate file
        use_new = tk.Button(self.review_window, text="Use Modified File", fg="green", bg='#acacad',
                        command = self.use_intermediate_file)
        use_new.pack(side="right")

    def cleaned_message(self):
        save_path = self.dirname + '/cleaned_data'
        self.review_window.destroy()
        self.cleaned_message = tk.Tk()
        self.cleaned_message.title(self.titled)
        note = tk.Label(self.cleaned_message)
        note["text"] = "Data import and cleaning is complete. \n" \
                       "You can find you cleaned, concatenated file here: \n\n" \
                        + save_path +\
                       "\n\n You can exit by clicking QUIT or \n" \
                       "run a standard PDF report of the hospital capacity data"

        note.pack(side="top")
        done = tk.Button(self.cleaned_message, text="QUIT", fg="red", bg='#acacad',
                         command=self.cleaned_message.destroy)
        done.pack(side="left")
        run_report = tk.Button(self.cleaned_message,
                               text="RUN PDF REPORT",
                               fg="green",
                               bg='#acacad',
                               command=self.run_pdf_report)
        run_report.pack(side="right")


    def run_pdf_report(self):
        """Create data plots and compile pdf report"""
        export_path = self.dirname + '/data_reports/'
        if not os.path.exists(export_path):
            os.makedirs(export_path)
        report_date = date.today().strftime("%B_%d_%Y")
        # Create a data segment with only the most recent data
        today_date = self.data['date'].max()
        today_data = self.hospitals[self.hospitals['date']==today_date]

        # Create plots
        report.create_bar_plot(today_data['HospitalName'],
                               today_data['TotalInpatientCOVIDDiagnosed'],
                        "Current Number of COVID Diagnoses InPatients - " + today_date, 'Hospital',
                        'Number of Inpatients',
                        export_path, "daily_COVID_inpatient_by_hospital", report_date
                        )

        report.create_bar_plot(today_data['HospitalName'],
                               today_data['TotalCOVIDDiagnosedOnVent'],
                               "Current Number of COVID-19 Patients on Vent - " + today_date, 'Hospital', "Number of Patients",
                               export_path, "daily_vent_use_by_hospital", report_date
                               )

        total_admits = self.hospitals.groupby(['date'])['YesterdayCOVIDAdmit'].sum()
        report.create_line_graph(total_admits,
                                 "Total Daily COVID Admissions All County Hospitals",
                                 "Date", "Number of Patients", export_path, "30day_COVID_admissions_all_county", report_date
                                )

        total_diagnosed = self.hospitals.groupby(['date'])['TotalInpatientCOVIDDiagnosed'].sum()
        report.create_line_graph(total_diagnosed,
                                 "Total COVID Diagnosed All County Hospitals",
                                 "Date", "Number of Patients", export_path, "30day_COVID_inpatients_total_county", report_date
                                )

        total_on_vent = self.hospitals.groupby(['date'])['TotalCOVIDDiagnosedOnVent'].sum()
        report.create_line_graph(total_on_vent,
                                 "Total COVID Diagnosed Patients on Ventilators- County Hospitals",
                                 "Date", "Number of Patients", export_path, "30day_vent_use_total_county", report_date
                                )

        diagnosed_by_hospital = self.hospitals.groupby(['date', 'HospitalName']).sum()['TotalInpatientCOVIDDiagnosed']
        report.create_multiline_graph(diagnosed_by_hospital,
                                      "Total COVID Inpatients by Hospital", "Date",
                                      "Number of Patients",
                                      export_path, "30day_vent_use_by_hospital", report_date)

        report.create_twoline_graph(total_on_vent, total_diagnosed, "Patients On Vents", "All Diagnosed Patients",
                                    "COVID-19 Diagnosed inpatients using and not using ventilators", "Date",
                                    "Number of Patients",
                                    export_path, "30day_proportional_vent_use", report_date)

        data_check = self.data['source_file'].value_counts().rename_axis("Source Files").to_frame('Num Imported Rows')
        import_check = data_check.sort_values(by=['Source Files'])

        # Create and save PDF report
        # save FPDF() class into a variable pdf
        pdf = fpdf.FPDF()
        # set style parameters

        pdf.set_margins(12.5, 20.5)

        # Add a page
        pdf.add_page()
        # create a title cell
        pdf.set_font("Times", size=16)
        pdf.cell(200, 10, txt="County Hospital Capacity", ln=1, align='C')
        # add another cell
        pdf.set_font("Times", size=11)
        pdf.cell(200, 8, txt="Report Date: " + str(date.today().strftime("%B %d, %Y")), ln=2, align='C')
        pdf.set_font("Times", "B", size=14)
        pdf.cell(200, 8, txt="Hospital Capacity Overview", ln=1, align='L')
        pdf.set_font("Times", size=11)
        pdf.cell(200, 8,
                 txt="This report overviews Hospital Capacity for 9 County Hospitals:",
                 ln=1, align='L')
        for hospital in self.hospital_names:
            pdf.cell(200, 8, txt="    - {}".format(hospital), ln=1, align='L')  #

        pdf.cell(200, 8, txt="The most recent data included here is from:\n" + today_date, ln=1, align='L')
        pdf.cell(200, 8, txt="This report is divided into the following sections:", ln=1, align='L')
        pdf.cell(200, 8, txt="    - COVID-19 Diagnosed Inpatients", ln=1, align='L')
        pdf.cell(200, 8, txt="         - data fields include: YesterdayCOVIDAdmit, TotalInpatientCOVIDDiagnosed", ln=1,
                 align='L')
        pdf.cell(200, 8, txt="    - COVID-19 Related Equipment Use", ln=1, align='L')
        pdf.cell(200, 8, txt="         - data fields include: TotalCOVIDDiagnosedOnVent, TotalInpatientCOVIDDiagnosed",
                 ln=1, align='L')
        pdf.cell(200, 5, txt="    - Data Sources + Cleaning", ln=1, align='L')

        pdf.set_font("Times", "B", size=14)
        pdf.cell(200, 8, txt="COVID-19 Diagnosed Inpatients", ln=1, align='L')
        pdf.set_font("times", size=12)

        pdf.cell(200, 8, txt="Reported number of current inpatients diagnosed with COVID-19 by hospital",
                 ln=1, align='L')
        pdf.image("{}{}_{}.jpg".format(export_path, "daily_COVID_inpatient_by_hospital", report_date), w=150)

        pdf.cell(200, 8, txt="Historical (30 day) reported COVID-19 admissions across County Hospitals",
                 ln=1, align='L')
        pdf.image("{}{}_{}.jpg".format(export_path, "30day_COVID_admissions_all_county", report_date), w=150)

        pdf.cell(200, 8, txt="Historical (30 day) number of inpatients diagnosed with COVID-19 by hospital" + today_date,
                 ln=1, align='L')
        pdf.image("{}{}_{}.jpg".format(export_path, "30day_COVID_inpatients_total_county",
                                       report_date), w=150)

        pdf.set_font("Times", "B", size=14)
        pdf.cell(200, 8, txt="Equipment Usage", ln=1, align='L')
        pdf.set_font("times", size=12)

        pdf.cell(200, 5, txt="Reported number of current COVID-19 patients on ventilators by hospital", ln=1, align='L')
        pdf.image("{}{}_{}.jpg".format(export_path, "daily_vent_use_by_hospital", report_date), w=150)

        pdf.cell(200, 5, txt="Historical (30 day) total reported number of COVID-19 patients on ventilators", ln=1,
                 align='L')
        pdf.image("{}{}_{}.jpg".format(export_path, "30day_vent_use_total_county", report_date), w=150)

        pdf.cell(200, 5, txt="Historical (30 day) reported number of COVID-19 patients on ventilators by hospital",
                 ln=1, align='L')
        pdf.image("{}{}_{}.jpg".format(export_path, "30day_vent_use_by_hospital", report_date), w=150)

        pdf.cell(200, 5, txt="Historical (30 day) reported number of COVID-19 patients on ventilators compared with total inpatients",
                 ln=1, align='L')
        pdf.image("{}{}_{}.jpg".format(export_path, "30day_proportional_vent_use", report_date), w=150)

        pdf.set_font("Times", "B", size=14)
        pdf.cell(200, 8, txt="Data Sources and Cleaning", ln=1, align='L')
        pdf.set_font("times", size=12)
        pdf.cell(200, 8, txt="Raw data may include multiple uploads. Graphs only contain the most recent data upload for each day",
                 ln=1, align='L')
        pdf.cell(200, 8, txt="Number of records imported per data source: ")
        pdf.cell(200, 8, txt=str(import_check))

        # save the pdf with name .pdf
        pdf.output("{}{}_{}.pdf".format(export_path, "Hospital_Capacity_Report", report_date))

        # Pop up final window
        self.cleaned_message.destroy()
        self.final_window = tk.Tk()
        self.final_window.title(self.titled)
        note = tk.Label(self.final_window)
        note["text"] = "PDF report is complete. \n" \
                       "You can find you the pdf report and figures jpgs here: \n\n" \
                       + export_path + \
                       "\n\n You can exit by clicking QUIT"
        note.pack(side="top")
        done = tk.Button(self.final_window, text="QUIT", fg="red", bg='#acacad', command=self.final_window.destroy)
        done.pack(side="left")

# Eventual main function:
LARGE_FONT = ("Verdana", 12)
root = tk.Tk()
# Launch the welcome window
title = "Data Import Tool"
greet = "\nWelcome to the Hospital Capacity data import tool. \n\n " \
    "This first step allows you to select new data you wish to import.  \n\n" \
    "1. Move all the Hospital Capacity files \n" \
    "you want to import into a single folder \n\n" \
    "2. Ensure the folder has \n" \
    " 'read/write/execute' access permissions \n\n" \
    "3. Select 'CHOOSE FOLDER' and choose the data folder \n\n" \
    "Note: This tool writes the cleaned data into \n" \
    "a subfolder named `cleaned` in the selected folder.\n"

app = Application(title, greet, master=root)
app.mainloop()















