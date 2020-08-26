######################################
##  ------------------------------- ##
##     SAW Template Generator       ##
##  ------------------------------- ##
##     Written by: David Shaeffer   ##
##  ------------------------------- ##
## Last Edited on: 06-25-2020       ##
##  ------------------------------- ##
######################################

try:
    import Tkinter as tk
    from Tkinter import *
except ImportError:
    import tkinter as tk
    from tkinter import *
# import tkMessageBox
import tkinter.messagebox as messagebox
import os
## Import library to get use ID - for building block lookup##
import getpass
## This is neccessary to important modules from python scripts folder after pyinstaller bundles other lirbaries 
##import sys
##sys.path.insert(0, os.path.abspath(os.curdir + '\Python_Scripts'))
import dateEntry
from datetime import datetime, timedelta
import json
from wordxml import WordXML
from locationservices import LocationServices
#for epcn data file
# import filedialog
import tkinter.filedialog as filedialog
from lxml import etree as ET

## Functions ##

# Get users curre template version
def getTVersion():
    try:
        # get path to current working directory
        cwd = os.getcwd()
        dir_path = os.path.abspath(cwd + '\Configuration_Files')
        # print dir_path
        for filename in os.listdir(dir_path):
            if filename.endswith(".xml") and filename[:2]=='GP':
                return filename
    except Exception as e:
        print ("Failed to get template version!")
        print (e)

# Parse json and return user value
def getSettings(value):
    try:
        # get path to current working directory
        cwd = os.getcwd()
        dir_path = os.path.abspath(cwd + r'\Configuration_Files\usersettings.json')
        with open(dir_path) as json_data:
            d = json.load(json_data)
        return d[value]
    except Exception as e:
        print ("Failed to load settings!")
        print (e)

# updates settings field in json file
def updateJSON(field, value):
    try:
        # get path to current working directory
        cwd = os.getcwd()
        dir_path = os.path.abspath(cwd + r'\Configuration_Files\usersettings.json')
        with open(dir_path, "r") as jsonFile:
            data = json.load(jsonFile)
        data[field] = value
        with open(dir_path, "w") as jsonFile:
            json.dump(data, jsonFile)
    except Exception as e:
        print ("Failed to updated JSON file!")
        print (e)

# saves settings and closes setting frame
def saveSettings():
    try:
        USER.grid_remove()
        # update fields based on user settings file
        updateJSON("username", nameent.get()) 
        updateJSON("useremail", emailent.get())
        updateJSON("userfieldoffice", FONAMES.get())
        updateJSON("userphone", phonent.get())
        updateJSON("bbfolder", bbfolderent.get()) 
        # updateJSON("configfolder", configfolderent.get())
        print ("Settings Saved!")
    except Exception as e:
        print ("Failed to save the settings!")
        print (e)

# Get users building block template version 
def getBBlock():
    try:
        if getSettings("bbfolder") == "":
            dir_path = "C:/Users/" + getpass.getuser() +"/AppData/Roaming/Microsoft/Document Building Blocks/1033/15"
        else:     
            dir_path = getSettings("bbfolder")
        # print dir_path
        for filename in os.listdir(dir_path):
            if filename.endswith(".dotx") and filename[:3]=='SAW':
                return filename
    except Exception as e:
        print ("Failed to load building block!")
        print (e)
# Generates an AR upload form for ORM
def generateARUpload(ARTable, path):
    us_state_abbrev  = {
        'AK': 'Alaska',
        'AL': 'Alabama',
        'AR': 'Arkansas',
        'AS': 'American Samoa',
        'AZ': 'Arizona',
        'CA': 'California',
        'CO': 'Colorado',
        'CT': 'Connecticut',
        'DC': 'District of Columbia',
        'DE': 'Delaware',
        'FL': 'Florida',
        'GA': 'Georgia',
        'GU': 'Guam',
        'HI': 'Hawaii',
        'IA': 'Iowa',
        'ID': 'Idaho',
        'IL': 'Illinois',
        'IN': 'Indiana',
        'KS': 'Kansas',
        'KY': 'Kentucky',
        'LA': 'Louisiana',
        'MA': 'Massachusetts',
        'MD': 'Maryland',
        'ME': 'Maine',
        'MI': 'Michigan',
        'MN': 'Minnesota',
        'MO': 'Missouri',
        'MP': 'Northern Mariana Islands',
        'MS': 'Mississippi',
        'MT': 'Montana',
        'NA': 'National',
        'NC': 'North Carolina',
        'ND': 'North Dakota',
        'NE': 'Nebraska',
        'NH': 'New Hampshire',
        'NJ': 'New Jersey',
        'NM': 'New Mexico',
        'NV': 'Nevada',
        'NY': 'New York',
        'OH': 'Ohio',
        'OK': 'Oklahoma',
        'OR': 'Oregon',
        'PA': 'Pennsylvania',
        'PR': 'Puerto Rico',
        'RI': 'Rhode Island',
        'SC': 'South Carolina',
        'SD': 'South Dakota',
        'TN': 'Tennessee',
        'TX': 'Texas',
        'UT': 'Utah',
        'VA': 'Virginia',
        'VI': 'Virgin Islands',
        'VT': 'Vermont',
        'WA': 'Washington',
        'WI': 'Wisconsin',
        'WV': 'West Virginia',
        'WY': 'Wyoming'}

    try:
        import csv
        with open(path, 'w', newline='') as csvfile:
            filewriter = csv.writer(csvfile, delimiter=',',
                                    quotechar='"', quoting=csv.QUOTE_MINIMAL)
            filewriter.writerow(['Waters_Name', 'State', 'Cowardin_Code', 'HGM_Code', 'Meas_Type', 'Amount', 'Units', 'Waters_Type', 'Latitude', 'Longitude', 'Local_Waterway'])
            for data in ARTable:
                if data[4] == 'acre':
                    units = 'ACRE'
                    measuretype = 'Area'
                elif data[4] == 'square feet':
                    units = 'SQ_FT'
                    measuretype = 'Area'
                elif data[4] == 'feet':
                    units = 'FOOT'
                    measuretype = 'Linear'
                state = us_state_abbrev[data[7]].upper()
                if data[8] == 'Impound':
                    rap = 'IMPNDMNT'
                elif data[8] == 'Isolated':
                    rap = 'ISOLATE'
                elif data[8] == 'Upland':
                    rap = 'UPLAND'
                elif data[8] == ' ':
                    rap = 'DELINEATE'
                else:
                    rap = data[8]
                filewriter.writerow([data[0], state, data[6], '', measuretype, data[3], units, rap, data[1], data[2], ''])
    except Exception as e:
        print ("Failed to generate AR Upload csv file!")
        print (e)

def createARTable(datafile):
    try:
        print("Populating AR Table")
        ARTable = []
        rowcounter = -1 

        # # get the number of AR Tables
        artablecount = datafile.xpath('count(//ARTable)')
        artotalcount = datafile.xpath("count(//ARTable/*[contains(local-name(),'WatersName')])")

        """create a list with sublist for each AR"""
        for a in range(int(artotalcount)):
            ARTable.append([])

        """for each ar table convert the xml table from the pdf to a list of lists"""
        for tables in datafile.iter('ARTable'):
            """count the ARs on each page"""
            arcount = datafile.xpath("count(//P%s/ARTable/*[contains(local-name(),'WatersName')])" % tables.getparent().tag.split("P", 1)[1])
            """loop through the count on each AR page"""
            for i in range(int(arcount)):
                """for each table row"""
                colcounter = 0
                try: 
                    for rows in tables:
                        """if you're on the same row iterate through each row - V2.0.1 patch to deal with no rapanos"""
                        if "Row" + str(i+1) == "Row" + rows.tag.split("Row",1)[1] and rows.text.isspace() is False:
                            """when there the column counter is zero, you know you are on the next row so increment the row"""
                            if colcounter == 0:
                                rowcounter += 1
                            """count the columns"""
                            colcounter += 1
                            """append the row data to the appopriate list"""
                            ARTable[rowcounter].append(rows.text)
                except IndexError:
                                pass
        """If the rapanos field is not present then insert a dummy item into the list - V2.0.1 patch to deal with no rapanos"""
        for i in ARTable:
            if len(i) == 8:
                i.insert(6, ' ')

        """Sort the list so it can be inserted into the wordxml table function"""
        try:
            order = [8, 2, 3, 5, 4, 0, 1, 7, 6]
            for a in range(int(artotalcount)):
                ARTable[a] = [ARTable[a][i] for i in order]
        except Exception as e:
            print("Unable to sort aquatic resources!")
            print(e)
        return ARTable
    except Exception as e:
        print ("Failed to create AR Table!")
        print (e)
    """Sort the list so it can be inserted into the wordxml table function"""

    
## Classes ##
class MainApplication(tk.Frame):
    import configurations

    ### Removes unneeded sections ###
    def removesections(self, wordtemplate):
        if gpts.get() == 0:
            wordtemplate.removesection(0) 
        if jdonlyts.get() == 0:
            wordtemplate.removesection(1)
        if cpfurnts.get() == 0:
            wordtemplate.removesection(2)
        if scts.get() == 0:
            wordtemplate.removesection(3)
        if cfts.get() == 0:
            wordtemplate.removesection(4)
        if ddocts.get() == 0:
            wordtemplate.removesection(5)
        if mtfts.get() == 0:
            wordtemplate.removesection(6)
        if gpjdts.get() == 0:
            wordtemplate.removesection(7) 
        if applts.get() == 0:
            wordtemplate.removesection(8) 
        if rapts.get() == 0:
            wordtemplate.removesection(9)
        if pjdts.get() == 0:
            wordtemplate.removesection(10)
            wordtemplate.removesection(11)
            wordtemplate.removesection(12)  
        if novts.get() == 0 and noncompts.get() == 0 :
            wordtemplate.removesection(13) 
            wordtemplate.removesection(14)  
        if nprts.get() == 0:
            wordtemplate.removesection(15)

    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent.title("SAW Regulatory Template Generator v2.0.4 BETA")

    ##### Main Screen #####
        ##Variables##
        global authority
        authority = BooleanVar()
        authority.set(False)
        # for the request type
        global reqtypjd
        reqtypjd = BooleanVar()
        reqtypjd.set(False)
        global reqtypgp
        reqtypgp = BooleanVar()
        reqtypgp.set(False)
        global reqtypnpr 
        reqtypnpr = BooleanVar()
        reqtypnpr.set(False)
        global reqtypnov
        reqtypnov = BooleanVar()
        reqtypnov.set(False)
        global reqtypscp
        reqtypscp = BooleanVar()
        reqtypscp.set(False)
        global reqtypnoncomp
        reqtypnoncomp = BooleanVar()
        reqtypnoncomp.set(False)
        # for the authority type
        global authtyp10
        authtyp10 = BooleanVar()
        authtyp10.set(False)
        global authtyp404
        authtyp404 = BooleanVar()
        authtyp404.set(False)
        ##Frame Widgets##
        content = tk.Frame(root)
        global JD
        JD = tk.Frame(root)
        global AJD
        AJD = tk.Frame(root)
        global GP
        GP = tk.Frame(root)
        global ESA
        ESA = tk.Frame(root)
        global CR
        CR = tk.Frame(root)
        global TEMP
        TEMP = tk.Frame(root)
        global USER
        USER = tk.Frame(root)

        ##Label widgets###
        geninfolb = tk.Label(content, text="General Project Information:", font='Helvetica 9 bold')
        bblocklb = tk.Label(content, text= getBBlock(), font='Helvetica 8 bold')
        templateverlb = tk.Label(content, text= getTVersion(), font='Helvetica 8 bold')
        filenumlb1 = tk.Label(content, text="File Number:")
        latlb = tk.Label(content, text="Latitude (Y):")
        lonlb = tk.Label(content, text="Longitude (X):")
        regauthlb = tk.Label(content, text="Regulatory Authority:")
        reqtypelb = tk.Label(content, text="Request Type:")

        ##Entry Widgets##
        global filenument1
        filenument1 = tk.Entry(content)
        global latent
        latent = tk.Entry(content)
        global lonent
        lonent = tk.Entry(content)

        ##Checkbox Widgets##
        global reqtypejdcb
        reqtypejdcb = tk.Checkbutton(content, text="Jurisdictional Determination", variable=reqtypjd, command=lambda v='JD': self.cb(v))
        global reqtypegpcb
        reqtypegpcb = tk.Checkbutton(content, text="General Permit", variable=reqtypgp, command=lambda v='GP': self.cb(v))
        global reqtypenprcb
        reqtypenprcb = tk.Checkbutton(content, text="No Permit Required", variable=reqtypnpr, command=lambda v='NPR': self.cb(v))
        global reqtypenovcb 
        reqtypenovcb = tk.Checkbutton(content, text="Unauthorized Activity", variable=reqtypnov, command=lambda v='NOV': self.cb(v))
        global reqtypenoncompcb
        reqtypenoncompcb = tk.Checkbutton(content, text="Non-Compliance", variable=reqtypnoncomp, command=lambda v='NC': self.cb(v))
        # global reqtypescpcb
        # reqtypescpcb = tk.Checkbutton(content, text="Scoping", variable=reqtypscp, command=lambda v='SCP': self.cb(v))
        global regauth10
        regauth10 = tk.Checkbutton(content, text="Section 10", variable=authtyp10, command=lambda v='10': self.cb(v))
        global regauth404
        regauth404 = tk.Checkbutton(content, text="Section 404", variable=authtyp404, command=lambda v='404': self.cb(v))

        ##Button Widgets##
        #ok = tk.Button(content, text="Okay")

        ##Grid Layout##
        # place frames in the grid
        content.grid(column=0, row=0)
        JD.grid(column=0, row=1, sticky="W")
        AJD.grid(column=0, row=2, sticky="W")
        GP.grid(column=0, row=3, sticky="W")
        ESA.grid(column=0, row=4, sticky="W")
        CR.grid(column=0, row=5, sticky="W")
        TEMP.grid(column=0, row=6, sticky="W")
        USER.grid(column=0, row=7, sticky="W")
        geninfolb.grid(sticky="W", column=0, row=0, columnspan=3)
        # template versions
        bblocklb.grid(column=4, row=0, columnspan=2, sticky="E")
        templateverlb.grid(column=4, row=1, columnspan=3, sticky="E")
        # file number
        filenumlb1.grid(sticky="W", row=1)
        filenument1.grid(sticky="W", column=1, row=1)
        # coordinates
        latlb.grid(sticky="E", column=4, row=2)
        latent.grid(sticky="E", column=5, row=2)
        lonlb.grid(sticky="E", column=4, row=3)
        lonent.grid(sticky="E", column=5, row=3)
        # regulatory authority check boxes
        regauthlb.grid(column=0, row=2, sticky="W")
        regauth10.grid(column=1, row=2, sticky="W")
        regauth404.grid(column=2, row=2, sticky="W")
        # action type check boxes
        reqtypelb.grid(column=0, row=4, sticky="W")
        reqtypegpcb.grid(column=1, row=4, sticky="W")
        reqtypejdcb.grid(column=2, row=4, sticky="W")
        reqtypenprcb.grid(column=3, row=4, sticky="W")
        reqtypenovcb.grid(column=1, row=5, sticky="W")
        # reqtypescpcb.grid(column=3, row=5, sticky="W")
        reqtypenoncompcb.grid(column=2, row=5, sticky="W")

     ##### GP Screen #####
        ##Variables##
        # for the request type
        global compmit
        compmit = BooleanVar()
        compmit.set(False)
        global compmitPRM
        compmitPRM = BooleanVar()
        compmitPRM.set(False)
        global specialcond
        specialcond = BooleanVar()
        specialcond.set(False)
        global atf
        atf = BooleanVar()
        atf.set(False)
        global waiver
        waiver = BooleanVar()
        waiver.set(False)
        global rwaiver
        rwaiver = BooleanVar()
        rwaiver.set(False)
        global imp408
        imp408 = BooleanVar()
        imp408.set(False)
        global impws
        impws = BooleanVar()
        impws.set(False)
        global agencycord
        agencycord = BooleanVar()
        agencycord.set(False)
        global corpscord
        corpscord = BooleanVar()
        corpscord.set(False)
        global esaefhmsa
        esaefhmsa = BooleanVar()
        esaefhmsa.set(False)
        global tribal106
        tribal106 = BooleanVar()
        tribal106.set(False)
        #ESA variables
        global esapresent
        esapresent = BooleanVar()
        esapresent.set(False)
        global efhpresent
        efhpresent = BooleanVar()
        efhpresent.set(False)
        global msapresent
        msapresent = BooleanVar()
        msapresent.set(False)
        global oaesa
        oaesa = BooleanVar()
        oaesa.set(False)
        global oaefh
        oaefh = BooleanVar()
        oaefh.set(False)
        #106/Tribal variables
        global nhpapresent
        nhpapresent = BooleanVar()
        nhpapresent.set(False)
        global oanhpa
        oanhpa = BooleanVar()
        oanhpa.set(False)
        global gtog
        gtog = BooleanVar()
        gtog.set(False)

        ##Label Widgets##
        gpheadlb = tk.Label(GP, text="General Permit Checklist: ", font='Helvetica 9 bold')
        gptypelb = tk.Label(GP, text="Permit (RGP or NWP): ")
        # gptypeor = tk.Label(GP, text=" OR ")
        prjchklb = tk.Label(GP, text="Select ALL that apply:  ")
        ### Options Menu Widgets ###
        # this function resets the selection list for the other om if the other om is selected
        def gpnumber(val):
            if RGPPERMITNUM.get() == val:
                NWPPERMITNUM.set("Pick a NWP Number")
            if NWPPERMITNUM.get() == val:
                RGPPERMITNUM.set("Pick a RGP Number")
        global NWPPERMITNUM
        NWPPERMITNUM = StringVar(GP)
        NWPPERMITNUM.set("Pick a NWP Number") # default value
        om1 = OptionMenu(GP, NWPPERMITNUM, "NWP 1. Aids to Navigation", "NWP 2. Structures in Artificial Canals", "NWP 3. Maintenance", "NWP 4. Fish and Wildlife Harvesting, Enhancement, and Attraction Devices and Activities"
                       , "NWP 5. Scientific Measurement Devices", "NWP 6. Survey Activities", "NWP 7. Outfall Structures and Associated Intake Structures", "NWP 8. Oil and Gas Structures on the Outer Continental Shelf",
                        "NWP 9. Structures in Fleeting and Anchorage Areas", "NWP 10. Mooring Buoys", "NWP 11. Temporary Recreational Structures", "NWP 12. Utility Line Activities", "NWP 13. Bank Stabilization",
                        "NWP 14. Linear Transportation Projects", "NWP 15. U.S. Coast Guard Approved Bridges", "NWP 16. Return Water from Upland Contained Disposal Areas", "NWP 17. Hydropower Projects", "NWP 18. Minor Discharges",
                        "NWP 19. Minor Dredging", "NWP 20. Response Operations for Oil and Hazardous Substances", "NWP 21. Surface Coal Mining Activities", "NWP 22. Removal of Vessels", "NWP 23. Approved Categorical Exclusions",
                        "NWP 24. Indian Tribe or State Administered Section 404 Programs", "NWP 25. Structural Discharges", "NWP 27. Aquatic Habitat Restoration, Establishment and Enhancement Activities", "NWP 28. Modifications of Existing Marinas",
                        "NWP 29. Residential Developments", "NWP 30. Moist Soil Management for Wildlife", "NWP 31. Maintenance of Existing Flood Control Facilities", "NWP 32. Completed Enforcement Actions", "NWP 33. Temporary Construction, Access and Dewatering",
                        "NWP 34. Cranberry Production Activities", "NWP 35. Maintenance Dredging of Existing Basins", "NWP 36. Boat Ramps", "NWP 37. Emergency Watershed Protection and Rehabilitation", "NWP 38. Cleanup of Hazardous and Toxic Waste",
                        "NWP 39. Commercial and Institutional Developments", "NWP 40. Agricultural Activities", "NWP 41. Reshaping Existing Drainage Ditches", "NWP 42. Recreational Facilities", "NWP 43. Stormwater Management Facilities",
                        "NWP 44. Mining Activities", "NWP 45. Repair of Uplands Damaged by Discrete Events", "NWP 46. Discharges in Ditches", "NWP 48. Existing Commercial Shellfish Aquaculture Activities", "NWP 49. Coal Remining Activities",
                        "NWP 50. Underground Coal Mining Activities", "NWP 51. Land Based Renewable Energy Generation Facilities", "NWP 52. Water Based Renewable Energy Generation Pilot Projects", "NWP 53. Removal of Low-Head Dams",
                        "NWP 54. Living Shorelines", command=gpnumber)
        global RGPPERMITNUM 
        RGPPERMITNUM = StringVar(GP)
        RGPPERMITNUM.set("Pick a RGP Number") # default value
        om2 = OptionMenu(GP, RGPPERMITNUM, "RGP197800056 Piers, Docks, Boathouses", "RGP197800080 Bulkheads and Riprap", "RGP197800125 Boat Ramps & Associated Piers, Docks and Structures", "RGP198000048 Emergency Activities on Ocean Beaches",
                        "PGP198000291 CAMA (NC Coastal Area Management Act)", "RGP198200030 Work in Waters of Lakes and Reservoirs", "RGP198200079 Corps Reservoirs", "RGP198200277 Manmade Basins and Canals", "RGP198500194 Artificial Reefs",
                        "RGP199200297 Discharge of Material on State or Federal Owned Property", "RGP199602878 Dredge and Discharge into Federally Authorized Navigational Channels", "RGP201600163 Charlotte Storm Water Services", "RGP198200031 NC DOT Bridges Widening Projects, Interchange Improvements", command=gpnumber)

        ##Checkbox Widgets##
        global compmitcb
        compmitcb = tk.Checkbutton(GP, text="Compensatory Mitigation", variable=compmit, command=lambda v='COMPMIT': self.cb(v))
        global compmitprmcb
        compmitprmcb = tk.Checkbutton(GP, text="Permittee Responsible Mitigation", variable=compmitPRM, command=lambda v='COMPMITPRM': self.cb(v))
        global specialcondcb
        specialcondcb = tk.Checkbutton(GP, text="Special Conditions", variable=specialcond, command=lambda v='SPECON': self.cb(v))
        global atfcb
        atfcb = tk.Checkbutton(GP, text="After-the-fact Permit", variable=atf, command=lambda v='ATF': self.cb(v))
        global waivercb
        waivercb = tk.Checkbutton(GP, text="General Condition Waiver", variable=waiver, command=lambda v='WAIVER': self.cb(v))
        global rwaivercb
        rwaivercb = tk.Checkbutton(GP, text="Regional Condition Waiver", variable=rwaiver, command=lambda v='RWAIVER': self.cb(v))
        global imp408cb
        imp408cb = tk.Checkbutton(GP, text="Affect to 408", variable=imp408, command=lambda v='IMP408': self.cb(v))
        global impwscb
        impwscb = tk.Checkbutton(GP, text="Affect to Wild and Scenic River", variable=impws, command=lambda v='IMPWS': self.cb(v))
        global agencycordcb
        agencycordcb = tk.Checkbutton(GP, text="Agency Coordination", variable=agencycord, command=lambda v='AGCORD': self.cb(v))
        global corpscordcb
        corpscordcb = tk.Checkbutton(GP, text="Corps Internal Coordination", variable=corpscord, command=lambda v='CRPCORD': self.cb(v))
        global esaefhmsacb
        esaefhmsacb = tk.Checkbutton(GP, text="Affect to ESA/EFH/MSA", variable=esaefhmsa, command=lambda v='ESAEFHMSA': self.cb(v))
        global tribal106cb
        tribal106cb = tk.Checkbutton(GP, text="Potential Affect to Tribal/106", variable=tribal106, command=lambda v='TRIB106': self.cb(v))
        
        ## Grid Layout ##
        gpheadlb.grid(column=0, row=0, columnspan=2, sticky="W")
        gptypelb.grid(column=0, row=1, columnspan=2, sticky="W")
        om1.grid(column=1, row=1, columnspan=5, sticky="W")
        # gptypeor.grid(column=2, row=1, sticky="W")
        om2.grid(column=1, row=2, columnspan=5, sticky="W")
        prjchklb.grid(column=0, row=3, sticky="W")
        # general checklist boxes
        compmitcb.grid(column=0, row=4, sticky="W")
        compmitprmcb.grid(column=1, row=4, sticky="W")
        specialcondcb.grid(column=2, row=6, sticky="W")
        waivercb.grid(column=2, row=4, sticky="W")
        rwaivercb.grid(column=3, row=4, sticky="W")
        imp408cb.grid(column=3, row=6, sticky="W")
        impwscb.grid(column=0, row=6, sticky="W")
        agencycordcb.grid(column=0, row=5, sticky="W")
        atfcb.grid(column=1, row=6, sticky="W")
        corpscordcb.grid(column=1, row=5, sticky="W")
        esaefhmsacb.grid(column=2, row=5, sticky="W")
        tribal106cb.grid(column=3, row=5, sticky="W")

        # hide the JD screen
        GP.grid_remove()

      ##### ESA/EFH/MSA Subscreen #####
        #section label
        esaseclb = tk.Label(ESA, text="For ESA/EFH/MSA Select ALL that Apply to the ACTION AREA:")
        
        # check boxes
        esapresentcb = tk.Checkbutton(ESA, text="ESA Species/Critical Habitat Present", variable=esapresent, command=lambda v='ESAPRESENT': self.cb(v))
        efhpresentcb = tk.Checkbutton(ESA, text="Essential Fish Habitat Present", variable=efhpresent, command=lambda v='EFHPRESENT': self.cb(v))
        msapresentcb = tk.Checkbutton(ESA, text="Magnuson-Stevens Act Review Required", variable=msapresent, command=lambda v='MSAPRESENT': self.cb(v))
        oaesacb = tk.Checkbutton(ESA, text="Other Agency Conducted ESA Review", variable=oaesa, command=lambda v='OAESA': self.cb(v))
        oaefhcb = tk.Checkbutton(ESA, text="Other Agency Conducted EFH Review", variable=oaefh, command=lambda v='OAEFH': self.cb(v))

        ##Grid Layout##
        # ESA section label
        esaseclb.grid(column=0, row=0, sticky="W", columnspan=4)
        # checkboxes
        esapresentcb.grid(column=0, row=1, sticky="W")
        efhpresentcb.grid(column=1, row=1, sticky="W")
        msapresentcb.grid(column=2, row=1, sticky="W")
        oaesacb.grid(column=0, row=2, sticky="W")
        oaefhcb.grid(column=1, row=2, sticky="W")

        # hide the esa screen
        ESA.grid_remove()

      ##### Tribal/106 Subscreen #####
        # section label
        crseclb = tk.Label(CR, text="For Tribal/NHPA 106 Select ALL that Apply to the PERMIT AREA:")

        # check boxes
        nhpapresentcb = tk.Checkbutton(CR, text="NHPA Section 106 Sites Present", variable=nhpapresent, command=lambda v='NHPAPRESENT': self.cb(v))
        oanhpacb = tk.Checkbutton(CR, text="Other Agency Conducted NHPA Review", variable=oanhpa, command=lambda v='OANHPA': self.cb(v))
        gtogcb = tk.Checkbutton(CR, text="Corps Conducted Government to Government Tribal Consultation", variable=gtog, command=lambda v='GTOG': self.cb(v))

        ##Grid Layout##
        # Tribal/106 section label
        crseclb.grid(column=0, row=0, sticky="W", columnspan=4)
        # checkboxes
        nhpapresentcb.grid(column=0, row=1, sticky="W")
        oanhpacb.grid(column=1, row=1, sticky="W")
        gtogcb.grid(column=2, row=1, sticky="W")

        # hide the CR screen
        CR.grid_remove()

    ##### JD Screen #####
        ##Variables##
        # JD type
        global typepjd
        typepjd = BooleanVar()
        typepjd.set(False)
        global typeajd
        typeajd= BooleanVar()
        typeajd.set(False)
        # JD checklist variables
        global uppresent
        uppresent = BooleanVar()
        uppresent.set(False)
        global wetpresent
        wetpresent = BooleanVar()
        wetpresent.set(False)
        # global validjd
        # validjd = BooleanVar()
        # validjd.set(False)

        # JD Checkbutton variables
        global typepjdcb
        global typeajdcb
        global uppresentcb
        global wetpresentcb
        global waterstypTNWcb 
        global waterstypTNWWcb 
        global waterstypeRPWcb 
        global waterstypTNWNRPWcb 
        global waterstypeRPWWDcb 
        global waterstypeRPWWNcb 
        global waterstypNRPWWNcb 
        global waterstypeIMPDcb
        global waterstypeISOLcb 
        global waterstypeNONJDcb 
        # global validjdcb

         # for the Rapanos AJD Water Types
        global waterstypTNW
        waterstypTNW = BooleanVar()
        waterstypTNW.set(False)
        global waterstypTNWW
        waterstypTNWW = BooleanVar()
        waterstypTNWW.set(False)
        global waterstypTNWRPW
        waterstypTNWRPW = BooleanVar()
        waterstypTNWRPW.set(False)
        global waterstypTNWNRPW
        waterstypTNWNRPW = BooleanVar()
        waterstypTNWNRPW.set(False)
        global waterstypRPW
        waterstypRPW = BooleanVar()
        waterstypRPW.set(False)
        global waterstypRPWWD
        waterstypRPWWD = BooleanVar()
        waterstypRPWWD.set(False)
        global waterstypRPWWN
        waterstypRPWWN = BooleanVar()
        waterstypRPWWN.set(False)
        global waterstypIMPD
        waterstypIMPD = BooleanVar()
        waterstypIMPD.set(False)
        global waterstypISOL
        waterstypISOL = BooleanVar()
        waterstypISOL.set(False)
        global waterstypNONJD
        waterstypNONJD = BooleanVar()
        waterstypNONJD.set(False)

        ##Entry Widgets##
        global filenument2
        filenument2 = tk.Entry(JD)
        
        ##Label Widgets##
        jdheadlb = tk.Label(JD, text="Jurisdictional Determination Checklist:", font='Helvetica 9 bold')
        jdtypelb = tk.Label(JD, text="Jurisdictional Determination Type:")
        # filenumlb2 = tk.Label(JD, text="Previous File Number:")
        # prevjddatelb = tk.Label(JD, text="Previous JD Date:")
        ##Checkbox Widgets##
        # JD type boxes
        typepjdcb = tk.Checkbutton(JD, text="Preliminary", variable=typepjd, command=lambda v='PJD': self.cb(v))
        typeajdcb = tk.Checkbutton(JD, text="Approved", variable=typeajd, command=lambda v='AJD': self.cb(v))
        # JD checklist boxes
        uppresentcb = tk.Checkbutton(JD, text="Select if ONLY UPLANDS are present in the review area", variable=uppresent, command=lambda v='UPP': self.cb(v))
        wetpresentcb = tk.Checkbutton(JD, text="Select if jurisdictional WETLANDS are present in the review area", variable=wetpresent, command=lambda v='WETP': self.cb(v))
        # validjdcb = tk.Checkbutton(JD, text="Select if there is a valid JD issued under another file number", variable=validjd, command=lambda v='VJD': self.cb(v))
        ##Data Entry Widget##
        # previous JD date
        # global dentry
        # dentry = dateEntry.DateEntry(JD, font=('Helvetica', 10, NORMAL), border=0)

        ##Grid Layout##
        jdheadlb.grid(sticky="W", column=0, row=0, columnspan=4)
        # JD type boxes
        jdtypelb.grid(column=0, row=3, sticky="W")
        typepjdcb.grid(column=1, row=3, sticky="W")
        typeajdcb.grid(column=2, row=3, sticky="W")
        # jursdicitonal check list checkboxes
        uppresentcb.grid(column=0, row=5, sticky="W", columnspan=4)
        wetpresentcb.grid(column=0, row=6, sticky="W", columnspan=4)
        # validjdcb.grid(column=0, row=7, sticky="W", columnspan=4)
        # regauth404.grid(column=0, row=5, sticky="W")
        # file number
        # filenumlb2.grid(column=0, row=8, columnspan=2)
        # filenument2.grid(column=1, row=8, columnspan=2)
        # previous JD date
        # prevjddatelb.grid(column=3, row=8)
        # dentry.grid(column=4, row=8)

        # hide the JD screen
        JD.grid_remove()

      #####AJD Subscreen##### 

        # IMPD - Impoundment/Pond
        # RPW - Relativley Permanent Water - done
        # RPWWD - Wetlands directly abutting RPWs that flow directly or indirectly into TNWs (wetlands, not streams) - done
        # RPWWN - Wetlands adjacent to but not directly abutting RPWs that flow directly or indirectly into TNWs (wetlands, not streams)
        # TNW - Traditional Navigable Waters (streams, not wetlands) - done
        # TNWW - Wetlands adjacent to TNWs (wetlands, not streams) - done
        # TNWRPW - Tributary consisting of both RPWs and non-RPWs (streams, not wetlands) - done
        # Upland - Uplans - Needed?

        ##Checkbox Widgets##
        ajdwaterstypelb = tk.Label(AJD, text="Select ALL Onsite JURISDICTIONAL Water Types:")
        waterstypTNWcb = tk.Checkbutton(AJD, text="Traditionally Navigable Waters", variable=waterstypTNW, command=lambda v='RWATER': self.cb(v))
        waterstypTNWWcb = tk.Checkbutton(AJD, text="Wetland Adjacent to TNWs", variable=waterstypTNWW, command=lambda v='RWETLAND': self.cb(v))
        waterstypeRPWcb = tk.Checkbutton(AJD, text="Relativley Permanent Waters (RPWs)", variable=waterstypRPW, command=lambda v='RWATER': self.cb(v))
        waterstypTNWNRPWcb = tk.Checkbutton(AJD, text="Non-RPWs that flow to TNWs", variable=waterstypTNWNRPW, command=lambda v='RWATER': self.cb(v))
        waterstypeRPWWDcb = tk.Checkbutton(AJD, text="Wetlands Directly Abutting RPWs", variable=waterstypRPWWD, command=lambda v='RWETLAND': self.cb(v))
        waterstypeRPWWNcb = tk.Checkbutton(AJD, text="Wetlands Adjacent to RPWs", variable=waterstypRPWWN, command=lambda v='RWETLAND': self.cb(v))
        waterstypNRPWWNcb = tk.Checkbutton(AJD, text="Wetlands Adjacent to Non-RPWs", variable=waterstypTNWRPW, command=lambda v='RWETLAND': self.cb(v))
        waterstypeIMPDcb = tk.Checkbutton(AJD, text="Impoundments/Ponds/OpenWater", variable=waterstypIMPD, command=lambda v='RWATER': self.cb(v))
        waterstypeISOLcb = tk.Checkbutton(AJD, text="JURISDICTIONAL Isolated Waters", variable=waterstypISOL, command=lambda v='RWATER': self.cb(v))
        waterstypeNONJDcb = tk.Checkbutton(AJD, text="Select if there are NON-JURISDICTIONAL waters onsite (isolated waters, upland ponds, BMPs, etc.)?", variable=waterstypNONJD, onvalue=True)

        ##Grid Layout##
        # AJD section label
        ajdwaterstypelb.grid(column=1, row=9, sticky="W")
        # waters type check boxes
        waterstypTNWcb.grid(column=1, row=10, sticky="W")
        waterstypTNWWcb.grid(column=2, row=10, sticky="W")
        waterstypTNWNRPWcb.grid(column=3, row=10, sticky="W")
        waterstypNRPWWNcb.grid(column=1, row=11, sticky="W")
        waterstypeRPWcb.grid(column=2, row=11, sticky="W")
        waterstypeRPWWDcb.grid(column=3, row=11, sticky="W")
        waterstypeRPWWNcb.grid(column=1, row=12, sticky="W")
        waterstypeIMPDcb.grid(column=2, row=12, sticky="W")
        waterstypeISOLcb.grid(column=3, row=12, sticky="W")
        #non-JD waters 
        waterstypeNONJDcb.grid(column=0, row=13, sticky="W", columnspan=5)

        ## hide the AJD screen
        AJD.grid_remove()
    
    ##### Selected Tempaltates/Save/Settings Screen #####
        ##Varibales##
        global gpts
        gpts = BooleanVar()
        gpts.set(False)
        global gpjdts
        gpjdts = BooleanVar()
        gpjdts.set(False)
        global cpfurnts 
        cpfurnts = BooleanVar()
        cpfurnts.set(False)
        global scts
        scts = BooleanVar()
        scts.set(False)
        global cfts
        cfts = BooleanVar()
        cfts.set(False)
        global applts
        applts = BooleanVar()
        applts.set(False)
        global ddocts
        ddocts = BooleanVar()
        ddocts.set(False)
        global mtfts
        mtfts = BooleanVar()
        mtfts.set(False)
        global rapts
        rapts = BooleanVar()
        rapts.set(False)
        global pjdts
        pjdts = BooleanVar()
        pjdts.set(False)
        global jdonlyts
        jdonlyts = BooleanVar()
        jdonlyts.set(False)
        global novts
        novts = BooleanVar()
        novts.set(False)
        global noncompts
        noncompts = BooleanVar()
        noncompts.set(False)
        global nprts
        nprts = BooleanVar()
        nprts.set(False)
        # global scopingts
        # scopingts = BooleanVar()
        # scopingts.set(False)

        ##Label Widgets##
        tempheadlb = tk.Label(TEMP, text="Selected Templates Based of Users Input: ", font='Helvetica 9 bold')

        ##Checkbox Widgets##
        global gptscb
        gptscb = tk.Checkbutton(TEMP, text="GP Tearsheet", variable=gpts, command=lambda v='GPTS': self.cb(v))
        global gpjdtscb
        gpjdtscb = tk.Checkbutton(TEMP, text="GP JD Tearsheet", variable=gpjdts, command=lambda v='GPJDTS': self.cb(v))
        global cpfurntscb
        cpfurntscb = tk.Checkbutton(TEMP, text="Copies Furnished", variable=cpfurnts, command=lambda v='CPFURN': self.cb(v))
        global sctscb
        sctscb = tk.Checkbutton(TEMP, text="Special Conditions", variable=scts, command=lambda v='SCTS': self.cb(v))
        global cftscb
        cftscb = tk.Checkbutton(TEMP, text="Compliance Form", variable=cfts, command=lambda v='CFTS': self.cb(v))
        global appltscb
        appltscb = tk.Checkbutton(TEMP, text="Appeals Form", variable=applts, command=lambda v='APPLTS': self.cb(v))
        global ddoctscb
        ddoctscb = tk.Checkbutton(TEMP, text="GP MFR/Decision Document", variable=ddocts, command=lambda v='DDOCTS': self.cb(v))
        global mtftscb
        mtftscb = tk.Checkbutton(TEMP, text="Mitigation Transfer Form", variable=mtfts, command=lambda v='MTFTS': self.cb(v))
        global raptscb
        raptscb = tk.Checkbutton(TEMP, text="NWPR AJD Form", variable=rapts, command=lambda v='RAPTS': self.cb(v))
        global pjdtscb
        pjdtscb = tk.Checkbutton(TEMP, text="PJD Form (attachment A)", variable=pjdts, command=lambda v='PJDTS': self.cb(v))
        global jdonlytscb
        jdonlytscb = tk.Checkbutton(TEMP, text="JD Only Tearhseet", variable=jdonlyts, command=lambda v='JDONLYTS': self.cb(v))
        global novtscb
        novtscb = tk.Checkbutton(TEMP, text="Unauthorized Activity", variable=novts, command=lambda v='NOVTS': self.cb(v))
        global noncomptscb
        noncomptscb = tk.Checkbutton(TEMP, text="Non-compliance", variable=noncompts, command=lambda v='NONCOMPTS': self.cb(v))
        global nprtscb
        nprtscb = tk.Checkbutton(TEMP, text="No Permit Required", variable=nprts, command=lambda v='NPRTS': self.cb(v))
        # global scopingtscb
        # scopingtscb = tk.Checkbutton(TEMP, text="Scoping Response", variable=scopingts, command=lambda v='SCOPINGTS': self.cb(v))

        ##Button Widgets##
        settingsbutton = Button(TEMP, text="User Settings", command=self.loadSettings) 
        savebutton = Button(TEMP, text="Save Template", command=self.saveFunc) 
        loaddatabutton = Button(TEMP, text="Load ePCN Data File", command=self.loadData)
        loadJDdatabutton = Button(TEMP, text="Load JD Data File", command=self.loadJDData) 

        ##Date Entry Widget##
        #signature label
        signlb = tk.Label(TEMP, text="Signature Date:")
        global dentry2
        dentry2 = dateEntry.DateEntry(TEMP, font=('Helvetica', 10, NORMAL), border=0)

        ## Grid Layout ##
        tempheadlb.grid(column=0, row=0, columnspan=2, sticky="W")
        # selected template check boxes
        gptscb.grid(column=0, row=1, sticky="W")
        gpjdtscb.grid(column=1, row=1, sticky="W")
        cpfurntscb.grid(column=2, row=1, sticky="W")
        sctscb.grid(column=3, row=1, sticky="W")
        cftscb.grid(column=4, row=1, sticky="W")
        appltscb.grid(column=0, row=2, sticky="W")
        ddoctscb.grid(column=1, row=2, sticky="W")
        mtftscb.grid(column=2, row=2, sticky="W")
        raptscb.grid(column=3, row=2, sticky="W")
        pjdtscb.grid(column=4, row=2, sticky="W")
        jdonlytscb.grid(column=0, row=3, sticky="W")
        novtscb.grid(column=1, row=3, sticky="W")
        noncomptscb.grid(column=2, row=3, sticky="W")
        nprtscb.grid(column=3, row=3, sticky="W")
        # scopingtscb.grid(column=4, row=3, sticky="W")
        #buttons
        # columnspan=2
        settingsbutton.grid(column=0, row=4, padx=5, pady=5, sticky="W")
        signlb.grid(column=3, row=4, columnspan=2)
        dentry2.grid(column=4, row=4, sticky="E")
        savebutton.grid(column=5, row=4, padx=5, pady=5)
        # loaddatabutton.grid(column=1, row=4, padx=5, pady=5)
        loadJDdatabutton.grid(column=2, row=4, padx=5, pady=5)
    
    ##### User Settings Screen #####

        ##Options Widget##
        global FONAMES 
        FONAMES= StringVar(GP)
        om3 = OptionMenu(USER, FONAMES, "Asheville", "Charlotte", "Raleigh", "Washington", "Wilmington")

        ##Label Widgets##
        usersettingslb = tk.Label(USER, text="User Settings: ", font='Helvetica 9 bold')
        namelb = tk.Label(USER, text="Project Manager Name:")
        emaillb = tk.Label(USER, text="Project Manager E-mail:")
        phonelb = tk.Label(USER, text="Project Manager Phone Number:")
        officelb = tk.Label(USER, text="Field Office Name:")
        bbfolderlb = tk.Label(USER, text="Users Microsoft Building Block Folder Path:")
        # configfolderlb = tk.Label(USER, text="Configurations Folder Server Path:")

        ##Entry Widgets##
        global nameent
        nameent = tk.Entry(USER)
        global emailent
        emailent = tk.Entry(USER)
        global phonent
        phonent = tk.Entry(USER)
        global bbfolderent
        bbfolderent = tk.Entry(USER)
        # global configfolderent
        # configfolderent = tk.Entry(USER)

        # populate entry fields to show current user settings
        nameent.delete(0, END)
        nameent.insert(0, getSettings("username"))
        emailent.delete(0, END)
        emailent.insert(0, getSettings("useremail"))
        phonent.delete(0, END)
        phonent.insert(0, getSettings("userphone"))
        if getSettings("bbfolder") == "":
            bbfolderent.delete(0, END)
            bbfolderent.insert(0, "C:/Users/" + getpass.getuser() +"/AppData/Roaming/Microsoft/Document Building Blocks/1033/15")
        else:     
            bbfolderent.delete(0, END)
            bbfolderent.insert(0, getSettings("bbfolder"))

        # if getSettings("configfolder") == "":
        #     configfolderent.delete(0, END)
        #     cwd = os.getcwd()
        #     configfolderent.insert(0, os.path.abspath(cwd + '\Configuration_Files'))
        # else:
        #     configfolderent.delete(0, END)
        #     configfolderent.insert(0, getSettings("configfolder"))

        # populate dropdown
        if getSettings("userfieldoffice")=="":
            FONAMES.set("Pick an office") # default value
        else: 
            FONAMES.set(getSettings("userfieldoffice"))

        ##Button Widget##
        savesettingsbut = Button(USER, text="Save/Close Settings", command=saveSettings)
        
        ##Options Menu Widget##

        ## Grid Layout ##
        usersettingslb.grid(column=0, row=0, columnspan=2, sticky="W")
        namelb.grid(column=0, row=1, sticky="W")
        nameent.grid(column=1, row=1, sticky="W")
        emaillb.grid(column=2, row=1, sticky="W")
        emailent.grid(column=3, row=1, sticky="W")
        phonelb.grid(column=0, row=2, sticky="W")
        phonent.grid(column=1, row=2, sticky="W")
        officelb.grid(column=2, row=2, sticky="W")
        om3.grid(column=3, row=2, sticky="W")
        bbfolderlb.grid(column=0, row=3, sticky="W")
        bbfolderent.grid(column=1, row=3, sticky="W", columnspan=5)
        # configfolderlb.grid(column=0, row=4, sticky="W")
        # configfolderent.grid(column=1, row=4, sticky="W", columnspan=5)
        savesettingsbut.grid(column=5, row=3, columnspan=2, sticky="E")
        
        ## hide the user settings screen
        USER.grid_remove()

    # saves template based on user input
    def saveFunc(self):
        if authtyp10.get() == False and authtyp404.get() ==False:
            authority.set(False); 
        ### Make sure needed fields are complete
        if (reqtypgp.get()==True or reqtypjd.get()==True or reqtypnov.get()==True or reqtypnoncomp.get()==True) and (authority.get()==False and uppresent.get()==False):
                messagebox.showinfo("Error", "Please select a jurisdictional authority!")
                print ("Please select a jurisdictional authority!")
        elif (reqtypjd.get()==True) and typepjd.get()==False and typeajd.get()==False:
             messagebox.showinfo("Error", "Please select a jurisdictional determination type!")
             print ("Please select a jurisdictional determination type!")
        elif (reqtypgp.get()==True) and NWPPERMITNUM.get() == "Pick a NWP Number" and RGPPERMITNUM.get() == "Pick a RGP Number": 
             messagebox.showinfo("Error", "Please select a permit number!")
             print ("Please select a permit number!")
        # elif gpts.get()==False and gpjdts.get()==False and cpfurnts.get()==False and scts.get()==False and cfts.get()==False and applts.get()==False and ddocts.get()==False and mtfts.get()==False and rapts.get()==False and pjdts.get()==False and jdonlyts.get()==False and novts.get()==False and noncompts.get()==False and noncompts.get()==False and nprts.get()==False and scopingts.get()==False: 
        elif gpts.get()==False and gpjdts.get()==False and cpfurnts.get()==False and scts.get()==False and cfts.get()==False and applts.get()==False and ddocts.get()==False and mtfts.get()==False and rapts.get()==False and pjdts.get()==False and jdonlyts.get()==False and novts.get()==False and noncompts.get()==False and noncompts.get()==False and nprts.get()==False: 
             messagebox.showinfo("Error", "No templates are selected for export!")
             print ("No templates are selected for export!")
        elif dentry2.get() == "":
            messagebox.showinfo("Error", "Please enter a valid date!")
            print ("Please enter a valid date!")
        else:
            ### Hide the GUI
            root.withdraw()
            ##Set the root location for the word xml document
                # get path to current working directory
            cwd = os.getcwd()
            dir_path = os.path.abspath(cwd + '\Configuration_Files')
            for filename in os.listdir(dir_path):
                if filename.endswith(".xml"):
                    if "GPJD" in filename: 
                        document = dir_path + "/" + os.path.join(filename)
                        wordtemplate = WordXML(document)
         
            #gets location data based on lat/long
            if latent.get() == "" or lonent.get() == "":
                messagebox.showinfo("WARNING", "Latitude/Longitude not complete! No location specific information will be provided!")
                print ("Latitude/Longitude not complete! No location specific information will be provided!")
            else:    
                """Get an instance of the LocationServices class"""  
                locationservice = LocationServices()      
                try:
                    #get google maps geocoding data based on lat/long
                    print ("Getting city, state, and county information..."),
                    locationinfo = locationservice.getProjectAreaInfo(latent.get(), lonent.get())
                    wordtemplate.editcdp('Town', str(locationinfo).split(',')[0])
                    wordtemplate.editcdp('State', str(locationinfo).split(',')[1])
                    wordtemplate.editcdp('County', str(locationinfo).split(',')[2])
                    print ("[Complete]")
                except:
                    print ("ERROR: Could not get local government information!")
                try:
                    #get data from EPA mywaters
                    print ("Conducting stream trace and getting HUC and nearest waterway..."),
                    watershed = locationservice.getWatershed(latent.get(), lonent.get())
                    wordtemplate.editcdp('HUC', str(watershed).split(',')[0])
                    wordtemplate.editcdp('Waterway', str(watershed).split(',')[1])
                    print ("[Complete]")
                except:
                    print ("ERROR: Could not get watershed information!")
                try:
                    #get quad from national map
                    print ("Getting quad information..."),
                    quad = locationservice.getQuad(latent.get(), lonent.get())
                    wordtemplate.editcdp('Quad', quad)
                    print ("[Complete]")
                except:
                    print ("ERROR: Could not get quad data!")
                try:
                    #get basin from NCOneMap
                    print ("Getting basin name..."),
                    basin = locationservice.getBasin(latent.get(), lonent.get())
                    wordtemplate.editcdp('Basin', basin)
                    print ("[Complete]")
                except:
                    print ("ERROR: Could not get basin data!")
                # try:
                #     if ddocts.get() == 1:
                #         #get species data
                #         print ("Getting IPAC species list..."),
                #         species = locationservice.getSpecies(str(locationinfo).split(',')[1], str(locationinfo).split(',')[2])
                #         wordtemplate.editcdp('Species', species)
                #         print ("[Complete]")
                #     else:
                #         pass
                # except:
                #     print ("ERROR: Could not get species data!") 
        
            print ("Populating template..."),
             ## check to see if ePCN data was loaded
            try: 
                filename.endswith(".xml") 
                #try parsing the epcn data
                try:
                    # setup parse to ignore special characters
                    parser = ET.XMLParser(encoding='utf-8', recover=True)
                    # parse the data file
                    datafile = ET.parse(filename, parser)

                    #try inserting data into form fields
                    try:
                        #insert property size
                        for item in datafile.iter('PropertySize'):
                            wordtemplate.editcdp('Size', item.text)

                        #insert agent information
                        for item in datafile.iter('agent'):
                            for child in item:
                                if child.tag == "Agent":
                                    wordtemplate.editcdp('AFirstName', child.text)
                                    agent = child.text
                                if child.tag == "AgBusinessName":
                                    wordtemplate.editcdp('ACompanyName', child.text)
                                if child.tag == "AgEmail":
                                    wordtemplate.editcdp('AEmail', child.text)
                                if child.tag == "AgTelephone":
                                    wordtemplate.editcdp('ATelephone', child.text)
                                if child.tag == "AgPostal":
                                    wordtemplate.editcdp('AZip', child.text)
                                if child.tag == "AgState":
                                    wordtemplate.editcdp('AState', child.text)
                                if child.tag == "AgCity":
                                    wordtemplate.editcdp('ACity', child.text)
                                if child.tag == "AgStreet1":
                                    wordtemplate.editcdp('AAddress1', child.text)
                                    address1=child.text
                                if child.tag == "AgStreet2":
                                    if child.text:
                                        wordtemplate.editcdp('AAddress1', address1 + ", " + child.text)
                        #check to see if there is an applicant, otherwise insert the agent into the permittee box
                        applicantisagent = 0
                        for item in datafile.iter('applicant'):
                            # if len(item.getchildren()) > 0:
                            for child in item:
                                #skip if agent is listed as applicant - this is rarely the accurate
                                if child.text == agent:
                                    applicantisagent = 1
                                    break
                                else:
                                    if child.tag == "ApplicantName":
                                        wordtemplate.editcdp('RFirstName', child.text)
                                    if child.tag == "ABusinessName":
                                        wordtemplate.editcdp('RCompanyName', child.text)
                                    if child.tag == "AEmail":
                                        wordtemplate.editcdp('REmail', child.text)
                                    if child.tag == "ATelephone":
                                        wordtemplate.editcdp('RTelephone', child.text)
                                    if child.tag == "APostal":
                                        wordtemplate.editcdp('RZip', child.text)
                                    if child.tag == "AState":
                                        wordtemplate.editcdp('RState', child.text)
                                    if child.tag == "ACity":
                                        wordtemplate.editcdp('RCity', child.text)
                                    if child.tag == "AStreet1":
                                        wordtemplate.editcdp('RAddress1', child.text)
                                        address1=child.text
                                    if child.tag == "AStreet2":
                                        if child.text:
                                            wordtemplate.editcdp('RAddress1', address1 + ", " + child.text)
                        #if applicant is agent then inset owner into applicant fields
                        if applicantisagent == 1:
                            for item in datafile.iter('OwnerInfo'):
                                for child in item:
                                    if child.tag == "OResponsibleParty":
                                        wordtemplate.editcdp('RFirstName', child.text)
                                    if child.tag == "Owner":
                                        wordtemplate.editcdp('RCompanyName', child.text)
                                    if child.tag == "OEmail":
                                        wordtemplate.editcdp('REmail', child.text)
                                    if child.tag == "OTelephone":
                                        wordtemplate.editcdp('RTelephone', child.text)
                                    if child.tag == "OPostal":
                                        wordtemplate.editcdp('RZip', child.text)
                                    if child.tag == "OState":
                                        wordtemplate.editcdp('RState', child.text)
                                    if child.tag == "OCity":
                                        wordtemplate.editcdp('RCity', child.text)
                                    if child.tag == "OStreet1":
                                        wordtemplate.editcdp('RAddress1', child.text)
                                        address1=child.text
                                    if child.tag == "OStreet2":
                                        if child.text:
                                            wordtemplate.editcdp('RAddress1', address1 + ", " + child.text)
                        
                        #insert owner info
                        for item in datafile.iter('OwnerInfo'):
                            for child in item:
                                if child.tag == "OResponsibleParty":
                                    wordtemplate.editcdp('OFirstName', child.text)
                                if child.tag == "Owner":
                                    wordtemplate.editcdp('OCompanyName', child.text)
                                if child.tag == "OEmail":
                                    wordtemplate.editcdp('OEmail', child.text)
                                if child.tag == "OTelephone":
                                    wordtemplate.editcdp('OTelephone', child.text)
                                if child.tag == "OPostal":
                                    wordtemplate.editcdp('OZip', child.text)
                                if child.tag == "OState":
                                    wordtemplate.editcdp('OState', child.text)
                                if child.tag == "OCity":
                                    wordtemplate.editcdp('OCity', child.text)
                                if child.tag == "OStreet1":
                                    wordtemplate.editcdp('OAddress1', child.text)
                                    address1=child.text
                                if child.tag == "OStreet2":
                                    if child.text:
                                        wordtemplate.editcdp('OAddress1', address1 + ", " + child.text)
                    except Exception as e:
                        print ("Could not insert ePCN data!")
                        print (e)
                except Exception as e:
                    print ("Error1 parsing data file!")
                    print (e)
                # print datafile
            except Exception as e:
                pass
                print ("No epcn data file loaded.")
                # print (e)

              ## check to see if JD data was loaded
            try: 
                print(JDfilename)
                JDfilename
                #try parsing the epcn data
                try:
                    #load the JD data file
                    datafile = ET.parse(JDfilename).getroot()

                    #insert agent information
                    for data in datafile:
                        if data.tag == "AgentFirstName":
                            wordtemplate.editcdp('AFirstName', data.text)
                            agent = data.text
                        if data.tag == "AgentLastName":
                            wordtemplate.editcdp('ALastName', data.text)
                            agent = data.text
                        if data.tag == "AgentCompany":
                            wordtemplate.editcdp('ACompanyName', data.text)
                        if data.tag == "AgentEmail":
                            wordtemplate.editcdp('AEmail', data.text)
                        if data.tag == "AgentPhone":
                            wordtemplate.editcdp('ATelephone', data.text)
                        if data.tag == "AgentZip":
                            wordtemplate.editcdp('AZip', data.text)
                        if data.tag == "AgentState":
                            wordtemplate.editcdp('AState', data.text)
                        if data.tag == "AgentCity":
                            wordtemplate.editcdp('ACity', data.text)
                        if data.tag == "AgentAddress":
                            wordtemplate.editcdp('AAddress1', data.text)
                        #insert property size
                        if data.tag == "Acreage":
                            wordtemplate.editcdp('Size', data.text)
                        #insert project name
                        if data.tag == "SiteName":
                            wordtemplate.editcdp('ProjectName', data.text)
                        #insert requestor info
                        if data.tag == "RequestorFirstName":
                            wordtemplate.editcdp('RFirstName', data.text)
                        if data.tag == "RequestorLastName":
                            wordtemplate.editcdp('RLastName', data.text)
                        if data.tag == "RequestorCompany":
                            wordtemplate.editcdp('RCompanyName', data.text)
                        if data.tag == "RequestorEmail":
                            wordtemplate.editcdp('REmail', data.text)
                        if data.tag == "RequestorPhone":
                            wordtemplate.editcdp('RTelephone', data.text)
                        if data.tag == "RequestorZip":
                            wordtemplate.editcdp('RZip', data.text)
                        if data.tag == "RequestorState":
                            wordtemplate.editcdp('RState', data.text)
                        if data.tag == "RequestorCity":
                            wordtemplate.editcdp('RCity', data.text)
                        if data.tag == "RequestorAddress":
                            wordtemplate.editcdp('RAddress1', data.text)
                        #insert owner info
                        if data.tag == "OwnerFirstName":
                            wordtemplate.editcdp('OFirstName', data.text)
                        if data.tag == "OwnerLastName":
                            wordtemplate.editcdp('OLastName', data.text)
                        if data.tag == "OwnerCompany":
                            wordtemplate.editcdp('OCompanyName', data.text)
                        if data.tag == "OwnerEmail":
                            wordtemplate.editcdp('OEmail', data.text)
                        if data.tag == "OwnerPhone":
                            wordtemplate.editcdp('OTelephone', data.text)
                        if data.tag == "OwnerZip":
                            wordtemplate.editcdp('OZip', data.text)
                        if data.tag == "OwnerState":
                            wordtemplate.editcdp('OState', data.text)
                        if data.tag == "OwnerCity":
                            wordtemplate.editcdp('OCity', data.text)
                        if data.tag == "OwnerAddress":
                            wordtemplate.editcdp('OAddress1', data.text)
                        if data.tag == "PIN1":
                            pins = data.text
                        if data.tag == "PIN2":
                            pins = pins + ", " + data.text
                        if data.tag == "PIN3":
                            pins = pins  + ", " + data.text
                        #insert data citations and check boxes
                        #JD Map
                        if data.tag == "DelineationMap" :
                            wordtemplate.editcdp('DelineationMap', data.text)
                            wordtemplate.editcdp('PJDMaps', 'true')
                            wordtemplate.editcdp('RPMaps', 'true')
                        if data.tag == "WetlandSheets" :
                            wordtemplate.editcdp('WetlandSheets', data.text)
                            wordtemplate.editcdp('PJDDataSheets', 'true')
                            wordtemplate.editcdp('RPDataSheets', 'true')
                        if data.tag == "PhotoLog" :
                            wordtemplate.editcdp('PhotoLog', data.text)
                            wordtemplate.editcdp('PJDPhotos', 'true')
                            wordtemplate.editcdp('RPPhotos', 'true')
                            wordtemplate.editcdp('RPPhotosother', 'true')
                            wordtemplate.editcdp('PJDPhotosother', 'true')
                        if data.tag == "FEMA" :
                            wordtemplate.editcdp('FEMA', data.text)
                            wordtemplate.editcdp('PJDFEMA', 'true')
                            wordtemplate.editcdp('RPFEMA', 'true')
                        if data.tag == "Topo" :
                            wordtemplate.editcdp('Topo', data.text)
                            wordtemplate.editcdp('PJDTopo', 'true')
                            wordtemplate.editcdp('RPTopo', 'true')
                        if data.tag == "NHD" :
                            wordtemplate.editcdp('NHD', data.text)
                            wordtemplate.editcdp('PJDHyrdoNHD', 'true')
                            wordtemplate.editcdp('RPHyrdoNHD', 'true')
                            wordtemplate.editcdp('RPHydro', 'true')
                            wordtemplate.editcdp('PJDHydro', 'true')
                        if data.tag == "NRCS" :
                            wordtemplate.editcdp('NRCS', data.text)
                            wordtemplate.editcdp('PJDSoils', 'true')
                            wordtemplate.editcdp('RPSoils', 'true')
                        if data.tag == "HUC" :
                            wordtemplate.editcdp('WBD', data.text)
                            wordtemplate.editcdp('PJDHyrdoWBD', 'true')
                            wordtemplate.editcdp('RPHyrdoWBD', 'true')
                            wordtemplate.editcdp('RPHydro', 'true')
                            wordtemplate.editcdp('PJDHydro', 'true')
                        if data.tag == "NWI" :
                            wordtemplate.editcdp('NWI', data.text)
                            wordtemplate.editcdp('PJDNWI', 'true')
                            wordtemplate.editcdp('RPNWI', 'true')
                        if data.tag == "WatersStudy" :
                            wordtemplate.editcdp('WatersStudy', data.text)
                            wordtemplate.editcdp('PJDNavStudy', 'true')
                            wordtemplate.editcdp('RPNavStudy', 'true')
                        if data.tag == "Imagery" :
                            wordtemplate.editcdp('Imagery', data.text)
                            wordtemplate.editcdp('PJDPhotosaerial', 'true')
                            wordtemplate.editcdp('RPPhotosaerial', 'true')
                            wordtemplate.editcdp('RPPhotosother', 'true')
                            wordtemplate.editcdp('PJDPhotosother', 'true')
                        if data.tag == "OtherSources" :
                            wordtemplate.editcdp('OtherSources', data.text)
                            wordtemplate.editcdp('PJDOther', 'true')
                            wordtemplate.editcdp('RPOther', 'true')
                    #complete location field with pin numbers
                    wordtemplate.editcdp('LocationDesc', pins)
                    #complete the PJD Table
                    try:
                        ARTable = createARTable(datafile)
                            
                        # Now populate the table with the list data    
                        wordtemplate.addtotable("PJDTable", ARTable)                

                    except Exception as e:
                        print ("Cannot populate PJD table.")

                    # This asks the user if they would like to generate an AR table
                    try:
                        result = messagebox.askyesno("Template Generator","Would you like to generate an ORM2 Aquatic Resource upload table?")
                        if result == True:
                            tablename = filedialog.asksaveasfilename(defaultextension=".csv",title = "Save ORM2 AR Upload Table",filetypes = (("csv files","*.csv"),("all files","*.*")))
                            generateARUpload(ARTable, tablename)
                            """Use the OS to open the xml file"""
                            os.startfile(tablename)
                        else:
                            pass
                    except Exception as e:
                        print(e)       
                  
                except Exception as e:
                    print ("Error2 parsing data file!")
                    print (e)
              
            except Exception as e:
                pass
                print ("No epcn data file loaded.")
                # print (e)
                        
            #complete general information fields
            wordtemplate.editcdp('DANumber', filenument1.get())
            wordtemplate.editcdp('Latitude', latent.get())
            wordtemplate.editcdp('Longitude', lonent.get())

            if NWPPERMITNUM.get()[:4] == 'NWP ':
                #set expiration date          
                wordtemplate.editcdp('ExpireDate', '03/18/2022')
                wordtemplate.editcdp('PermitNumber', NWPPERMITNUM.get())
            elif RGPPERMITNUM.get()[:4] == 'RGP1':
                #set expiration date          
                wordtemplate.editcdp('ExpireDate', '12/31/2021')
                wordtemplate.editcdp('PermitNumber', RGPPERMITNUM.get())
            elif RGPPERMITNUM.get()[:4] == 'RGP2':
                #set expiration date          
                wordtemplate.editcdp('ExpireDate', '04/24/2022')
            
            #set JD date as signed date           
            JDDate = datetime.strptime(dentry2.get(), "%d%m%Y").date()
            wordtemplate.editcdp('JDDate', JDDate.strftime('%m/%d/%Y'))
            #set signed date           
            JDDate = datetime.strptime(dentry2.get(), "%d%m%Y").date()
            wordtemplate.editcdp('SignDate', JDDate.strftime('%m/%d/%Y'))

            ### JD Functions ###
            #if the JD type selected is prelimianry and uplands are not present
            if typepjd.get() == 1 and uppresent.get()==0:
                #indicate waters are present
                wordtemplate.updatedropdown('Waters', 1)

                #Check the appeal form
                wordtemplate.editcdp('AF_PJD', 'true')

                #Select appropirate JD type on JD letter
                wordtemplate.editcdp('PJDCheck', 'true')
                wordtemplate.editcdp('PJD1Check', 'true')

                #Specify PJD in JD basis
                wordtemplate.editcdp('JDType', 'preliminary')
                
                #Set appeal date to N/A
                wordtemplate.editcdp('AppealDate', 'Not applicable')

                #Set expire date to N/A
                wordtemplate.editcdp('JDExpireDate', 'Not applicable')

            # if the JD type slected is approved do this
            if typeajd.get()==1:
                #Check appeal form
                wordtemplate.editcdp('AF_AJD', 'true')
                if uppresent.get() != True or (authtyp10.get() != False or authtyp404.get()!=False):
                    wordtemplate.editcdp('AJD2Check', 'true')
                
                #specify AJD in JD basis
                wordtemplate.editcdp('JDType', 'approved')
                            
                #Insert appeal date            
                date_60_days = datetime.strptime(dentry2.get(), "%d%m%Y").date() + timedelta(days=60)
                wordtemplate.editcdp('AppealDate', date_60_days.strftime('%m/%d/%Y'))

                #Add expiration date
                date_5_years = datetime.strptime(dentry2.get(), "%d%m%Y").date() + timedelta(days=1825)
                wordtemplate.editcdp('JDExpireDate', date_5_years.strftime('%m/%d/%Y'))
                                
            #select appropriate JD type on JD letter
            if authtyp404.get()==1 and authtyp10.get()==1 and typeajd.get()==1:
                #select AJD Section 10 and 404 checkbox
                wordtemplate.editcdp('AJD10404Check', 'true')
                wordtemplate.updatedropdown('Waters', 1)
            if authtyp404.get()==1 and typeajd.get()==1:
                #select AJD Section 404 checkbox
                wordtemplate.editcdp('AJD404Check', 'true')
                wordtemplate.updatedropdown('Waters', 1)
            if wetpresent.get() == 1:
                wordtemplate.updatedropdown('Waters', 2)

            #hanlde previous JD - removed due to lack of value
            # if validjd.get() == 1:
            #     #insert previous file number
            #     wordtemplate.editSDP('PreviousJDFileNum', filenument2.get())
            #     #set JD issued date
            #     date = datetime.strptime(dentry.get(), "%d%m%Y").date()
            #     wordtemplate.editSDP('PreviousJDIssueDate', date.strftime('%m/%d/%Y'))
            #     #check appropriate box on JD form
            #     wordtemplate.editcdp('PreviousJDCheck', 1)
            #     #set appeal date to N/A
            #     wordtemplate.editcdp('AppealDate', 'Not applicable.')
            #     #set expire date to see preivous JD
            #     wordtemplate.editcdp('JDExpireDate', 'See previously issued JD')
            #     #set JD signed date to previous JD
            #     wordtemplate.editcdp('JDDate', date.strftime('%m/%d/%Y'))
            #     #set basis language
            #     wordtemplate.editcdp('JDType', 'previously issued preliminary or approved')
                
            #### Select applicable law(s) ###
            # if only 404 check 404 box and update section A and B on rapanos  
            if authtyp404.get()==1:
                wordtemplate.editcdp('Section404_CB', 'true')
                wordtemplate.updatedropdown('PermitAuthority', 2)
                wordtemplate.updatedropdown('RSection10', 2)
                wordtemplate.updatedropdown('RSection404', 1)
            # if only 10 check 10 box and update section A and B on rapanos
            if authtyp10.get()==1:
                wordtemplate.editcdp('Section10_CB', 'true')
                wordtemplate.updatedropdown('PermitAuthority', 1)
                wordtemplate.updatedropdown('RSection10', 1)
                wordtemplate.updatedropdown('RSection404', 2)
            #if 10 and 404 check both boxes and update section A and B on rapanos   
            if authtyp10.get()==1 and authtyp404.get()==1: 
                wordtemplate.editcdp('Section404_CB', 'true')
                wordtemplate.editcdp('Section10_CB', 'true')
                wordtemplate.updatedropdown('PermitAuthority', 3)
                wordtemplate.updatedropdown('RSection10', 1)
                wordtemplate.updatedropdown('RSection404', 1)

            ### Upland Only JD ###
            if uppresent.get()==1:
                wordtemplate.updatedropdown('RSection10', 2)
                wordtemplate.updatedropdown('RSection404', 2)
                wordtemplate.editcdp('NoWOUSCheck', 'true')   

            ### Enforcement ###
            if reqtypnov.get() == True:
                wordtemplate.updatedropdown('EnforcementType', 1) 
                wordtemplate.editcdp('Unauthorized_CB', 'true')
                wordtemplate.removeCC('Noncompliance')
            if reqtypnoncomp.get() == True:
                wordtemplate.updatedropdown('EnforcementType', 2)
                wordtemplate.editcdp('NonComp_CB', 'true')
                wordtemplate.removeCC('Unauthorized')

            ### Complete Applicable Sections of Decision Document ###
            #complete special conditions section
            if specialcond.get() == 1:
                wordtemplate.updatedropdown('SCPickList', 1)
                wordtemplate.removeCC('SCNO')
            else:
                wordtemplate.updatedropdown('SCPickList', 2)
                wordtemplate.removeCC('SCSection')
            #coordinated with external agency?
            if agencycord.get() == 1:
                wordtemplate.updatedropdown('CoordinatedAgency', 1)
            else:
                wordtemplate.updatedropdown('CoordinatedAgency', 2)
                wordtemplate.removeCC('CoordinateAgencySection')
            #is a stf verification
            if atf.get() == 1:
                wordtemplate.updatedropdown('ATF', 1)
            else:
                wordtemplate.updatedropdown('ATF', 2)

            #is a general waiver required
            if waiver.get() == 1:
                wordtemplate.updatedropdown('NWPWaiver', 1)
                # wordtemplate.updatedropdown('WaiverConclusion', 1)
            else:
                wordtemplate.updatedropdown('NWPWaiver', 2)
                # wordtemplate.updatedropdown('NWPWaiverConclusion', 2)
                wordtemplate.removeCC('NWPWaiverLang1')
                wordtemplate.removeCC('NWPWaiverLang2')
            #is a regional waiver required
            if rwaiver.get() == 1:
                wordtemplate.updatedropdown('RGPWaiver', 1)
                # wordtemplate.updatedropdown('RGPWaiverConclusion', 1)
            else:
                wordtemplate.updatedropdown('RGPWaiver', 2)
                # wordtemplate.updatedropdown('RGPWaiverConclusion', 2)
                wordtemplate.removeCC('RGPWaiverLang1')
                wordtemplate.removeCC('RGPWaiverLang2')
                wordtemplate.removeCC('RGPWaiverLang3')
                wordtemplate.removeCC('RGPWaiverRationale')
            #wild and scenic river?
            if impws.get() == 1:
                wordtemplate.updatedropdown('WSDropDown', 1)
            else:
                wordtemplate.updatedropdown('WSDropDown', 2)
                wordtemplate.removeCC('WSSection')
            #is compensatory mitigation permittee responsible?
            if compmitPRM.get() == 0:
                wordtemplate.removeCC('PermitteeResponsible')
            else:
                pass
            #is compensatory mitigation required?
            if compmit.get() == 0:
                wordtemplate.removeCC('MitigationYes')
                wordtemplate.updatedropdown('MitigationRequired', 2)
                wordtemplate.updatedropdown('MitigationConclusion1', 1)
                wordtemplate.updatedropdown('MitigationConclusion2', 1)
            else:
                wordtemplate.updatedropdown('MitigationRequired', 1)
                wordtemplate.updatedropdown('MitigationConclusion1', 2)
                wordtemplate.updatedropdown('MitigationConclusion2', 2)
            #would the project impact a civil works project
            if imp408.get() == 0:
                wordtemplate.updatedropdown('408DropDown', 1)
                wordtemplate.removeCC('408Section')
            else:
                wordtemplate.updatedropdown('408DropDown', 3)
            
            #was PCN coordinated with other Corps offices?
            if corpscord.get() == 1:
                wordtemplate.updatedropdown('CoordinateCorps', 1)
            else:
                wordtemplate.updatedropdown('CoordinateCorps', 2)
                wordtemplate.removeCC('CoordinateCorpsSection')
                
            ### ESA/EFH Sections of Decision Document ###
            #agency coordinate for ESA
            if oaesa.get() == 1:
                wordtemplate.updatedropdown('ESAOtherAgencyDropDown', 1)
            else:
                wordtemplate.updatedropdown('ESAOtherAgencyDropDown', 2)
                wordtemplate.removeCC('ESAOtherAgencySection')
             #known ESA species/critical habitat?
            if esapresent.get() == 1:
                wordtemplate.updatedropdown('ESAPresentDropDown', 1)
            else:
                wordtemplate.updatedropdown('ESAPresentDropDown', 2)
            #agency coordinate for EFH
            if oaefh.get() == 1:
                wordtemplate.updatedropdown('EFHOtherAgencyDropDown', 1)
            else:
                wordtemplate.updatedropdown('EFHOtherAgencyDropDown', 2)
                wordtemplate.removeCC('EFHOtherAgencySection')
            #require MSA review?
            if msapresent.get() == 1:
                wordtemplate.updatedropdown('MSADropDown', 1)
            else:
                wordtemplate.updatedropdown('MSADropDown', 2)
                wordtemplate.removeCC('EFHSection')
            #known EFH?
            if efhpresent.get() == 1:
                wordtemplate.updatedropdown('EFHPresentDropDown', 1)
            else:
                wordtemplate.updatedropdown('EFHPresentDropDown', 2)

            ### Section 106 NHPA & Tribal ###
            #agency coordinate for 106
            if oanhpa.get() == 1:
                wordtemplate.updatedropdown('OASection106DropDown', 1)
            else:
                wordtemplate.updatedropdown('OASection106DropDown', 2)
                wordtemplate.removeCC('OASection106Section')
            #know 106 properties?
            if nhpapresent.get() == 1:
                wordtemplate.updatedropdown('Section106DropDown', 1)
            else:
                wordtemplate.updatedropdown('Section106DropDown', 2)
            #kG to G coordination?
            if gtog.get() == 1:
                wordtemplate.updatedropdown('TribalTrustDropDown', 1)
            else:
                wordtemplate.updatedropdown('TribalTrustDropDown', 2)

            ## complete user specific sections
            if nameent.get() != "":
                wordtemplate.editcdp('PMName', nameent.get())
            if emailent.get() != "":
                wordtemplate.editcdp('PMEmail', emailent.get())
            if phonent.get() != "":
                wordtemplate.editcdp('PMPhone', phonent.get())
            if FONAMES.get() == "Charlotte":
                wordtemplate.updatedropdown('PMOffice', 2)
                wordtemplate.updatedropdown('PMOfficeAddress1', 2)
                wordtemplate.updatedropdown('PMOfficeAddress2', 2)
            elif FONAMES.get() == "Asheville":
                wordtemplate.updatedropdown('PMOffice', 3)
                wordtemplate.updatedropdown('PMOfficeAddress1', 3)
                wordtemplate.updatedropdown('PMOfficeAddress2', 3)
            elif FONAMES.get() == "Raleigh":
                wordtemplate.updatedropdown('PMOffice', 6)
                wordtemplate.updatedropdown('PMOfficeAddress1', 5)
                wordtemplate.updatedropdown('PMOfficeAddress2', 5)
            elif FONAMES.get() == "Washington":
                wordtemplate.updatedropdown('PMOffice', 4)
                wordtemplate.updatedropdown('PMOfficeAddress1', 6)
                wordtemplate.updatedropdown('PMOfficeAddress2', 6)
            elif FONAMES.get() == "Wilmington":
                wordtemplate.updatedropdown('PMOffice', 5)
                wordtemplate.updatedropdown('PMOfficeAddress1', 4)
                wordtemplate.updatedropdown('PMOfficeAddress2', 4)

            ##Select appopriate waters on the Rapanos Form and higlight sections user must complete on rapanos
            if waterstypTNW.get() == 1:
                wordtemplate.editcdp('RPTNWCheck', 'true')
                 # if TNW selected you must keep index 0, 11, and 6 to highlight section user neeed to complete
                wordtemplate.addHL(0)
                wordtemplate.addHL(1)
                wordtemplate.addHL(3)
                wordtemplate.addHL(4)
                wordtemplate.addHL(14)
                wordtemplate.addHL(15)
                wordtemplate.addHL(24)
            if waterstypTNWW.get() == 1:
                wordtemplate.editcdp('RPTNWWCheck', 'true')
                # if wetlands adjacent to TNW keep index 0, 1, 2, 11, and 12 to highlight section user neeed to complete
                wordtemplate.addHL(0)
                wordtemplate.addHL(1)
                wordtemplate.addHL(3)
                wordtemplate.addHL(4)
                wordtemplate.addHL(5)
                wordtemplate.addHL(14)
                wordtemplate.addHL(15)
                wordtemplate.addHL(24)
            if waterstypRPW.get() == 1:
                wordtemplate.editcdp('RPRPWCheck', 'true')
                # if RPW keep index 13 and 11 to highlight section user neeed to complete
                wordtemplate.addHL(0)
                wordtemplate.addHL(1)
                wordtemplate.addHL(14)
                wordtemplate.addHL(16)
                wordtemplate.addHL(24)
            if waterstypTNWNRPW.get() == 1:
                wordtemplate.editcdp('RPNONRPWCheck', 'true')
                 # if Non-RPWs that flow to TNWs keep index 8, 11, 14 to highlight section user neeed to complete
                wordtemplate.addHL(0)
                wordtemplate.addHL(1)
                wordtemplate.addHL(6)
                wordtemplate.addHL(7)
                wordtemplate.addHL(8)
                wordtemplate.addHL(9)
                wordtemplate.addHL(10)
                wordtemplate.addHL(11)
                wordtemplate.addHL(14)
                wordtemplate.addHL(17)
                wordtemplate.addHL(24)
            if waterstypRPWWD.get() == 1:
                wordtemplate.editcdp('RPRPWWDCheck', 'true')
                 # if Wetlands Directly Abutting RPWs keep index 11 and 15 to highlight section user neeed to complete
                wordtemplate.addHL(0)
                wordtemplate.addHL(1)
                wordtemplate.addHL(14)
                wordtemplate.addHL(18)
                wordtemplate.addHL(24)
            if waterstypRPWWN.get() == 1:
                wordtemplate.editcdp('RRPRPWWNCheck', 'true')
                 # if Wetlands Adjacent to RPWs keep index 3,4,5,6,10,16 to highlight section user neeed to complete
                wordtemplate.addHL(0)
                wordtemplate.addHL(1)
                wordtemplate.addHL(6)
                wordtemplate.addHL(7)
                wordtemplate.addHL(8)
                wordtemplate.addHL(9)
                wordtemplate.addHL(10)
                wordtemplate.addHL(13)
                wordtemplate.addHL(19)
                wordtemplate.addHL(24)
            if waterstypTNWRPW.get() == 1:
                wordtemplate.editcdp('RPNRPWWNCheck', 'true')
                 # if Wetlands Adjacent to Non-RPW keep index 3,4,5,6,10,16 to highlight section user neeed to complete
                wordtemplate.addHL(0)
                wordtemplate.addHL(1)
                wordtemplate.addHL(6)
                wordtemplate.addHL(7)
                wordtemplate.addHL(8)
                wordtemplate.addHL(9)
                wordtemplate.addHL(10)
                wordtemplate.addHL(12)
                wordtemplate.addHL(20)
                wordtemplate.addHL(24)
            if waterstypIMPD.get() == 1:
                wordtemplate.editcdp('RPIMPDCheck', 'true')
                wordtemplate.addHL(0)
                wordtemplate.addHL(1)
                wordtemplate.addHL(21)
                wordtemplate.addHL(24)
            if waterstypISOL.get() == 1:
                wordtemplate.editcdp('RPISOCheck', 'true')
                wordtemplate.addHL(0)
                wordtemplate.addHL(1)
                wordtemplate.addHL(22)
                wordtemplate.addHL(24)
            if waterstypNONJD.get() == 1:
                wordtemplate.editcdp('RPNONJDCheck', 'true')
                wordtemplate.addHL(0)
                wordtemplate.addHL(1)
                wordtemplate.addHL(23)
                wordtemplate.addHL(24)

            ### Removed unneeded sections based on check boxes ###
            self.removesections(wordtemplate)

            ##populating template complete
            print ("[Complete]")

            ### Save the xml file with updates based on user entry ###
            print ("Opening save dialog..."),
            wordtemplate.savefile()
            print ("[Complete]")
            ###Show the GUI in the case the PM needs to export the template again
            root.update()
            root.deiconify()          

    # loads data file
    def loadData(self):
        # get path to the data file
        global filename 
        filename = filedialog.askopenfilename() # show an "Open" dialog box and return the path to the selected file
    
     # loads data file
    def loadJDData(self):
        # get path to the data file
        global JDfilename 
        JDfilename = filedialog.askopenfilename() # show an "Open" dialog box and return the path to the selected file
        """"Fill in the GUI based on the JD form data"""
        #select JD since is a JD data file - TODO should it prompt to ask if this is a JD in the case it is a permit?!
        reqtypejdcb.select()
        JD.grid()
        #parse JD data file
        xmldata = ET.parse(JDfilename).getroot()
        # sort the AR tables
        ARTable = createARTable(xmldata)
        section10 = False
        section404 = False
        for row in ARTable:
            # print(row)
            if row[5] == '404':
                section404 = True
            elif row[5] =='10':
                section10 = True
            elif row[5] == '10/404':
                section10=True
            if "EM" in row[6] or "FO" in row[6] or "R6" in row[6] or "RP" in row[6] or "SS" in row[6]:
                wetpresentcb.select()
        if section404 == True:
            regauth404.select()
            authority.set(True);        
        if section10==True:
            regauth10.select()
            authority.set(True)
        # clear lat long then fill out sections of the user interface based on JD form data
        latent.delete(0, 'end')
        lonent.delete(0, 'end')
        for data in xmldata:
            # print(data)
            if data.tag == 'Latitude':
                latent.insert(0, str(data.text))
            if data.tag == 'Longitude':
                lonent.insert(0, str(data.text))
            if data.tag == 'Approved':
                if data.text == 'Yes':
                    typeajdcb.select()
                    #show AJD screen
                    #AJD.grid()
                   
                    #make sure PJD is not checked
                    typepjdcb.deselect()
                    # select needed tamples
                    cpfurntscb.select()
                    appltscb.select()
                    jdonlytscb.select()
                    #deselect unneeded templates
                    gptscb.deselect()
                    gpjdtscb.deselect()
                    ddoctscb.deselect()
                    cftscb.deselect()
                    # scopingtscb.deselect()
                    nprtscb.deselect()
                    noncomptscb.deselect()
                    #deselect GP otions
                    compmitcb.deselect()
                    compmitprmcb.deselect()
                    specialcondcb.deselect()
                    atfcb.deselect()
                    rwaivercb.deselect()
                    waivercb.deselect()
                    imp408cb.deselect()
                    impwscb.deselect()
                    agencycordcb.deselect()
                    corpscordcb.deselect()
                    esaefhmsacb.deselect()
                    tribal106cb.deselect()
                    #select rapanos template
                    raptscb.select() 
                if data.text == 'Off':
                    pass
            if data.tag == 'Preliminary':
                if data.text == 'Yes':
                    typepjdcb.select()
                    #make sure PJD is not checked
                    typeajdcb.deselect()
                    #remove ajd grid
                    AJD.grid_remove()
                    # select needed templates
                    cpfurntscb.select()
                    appltscb.select()
                    jdonlytscb.select()
                    #deselect unneeded templates
                    gptscb.deselect()
                    gpjdtscb.deselect()
                    ddoctscb.deselect()
                    cftscb.deselect()
                    # scopingtscb.deselect()
                    nprtscb.deselect()
                    noncomptscb.deselect()
                    #deselect GP otions
                    compmitcb.deselect()
                    compmitprmcb.deselect()
                    specialcondcb.deselect()
                    atfcb.deselect()
                    rwaivercb.deselect()
                    waivercb.deselect()
                    imp408cb.deselect()
                    impwscb.deselect()
                    agencycordcb.deselect()
                    corpscordcb.deselect()
                    esaefhmsacb.deselect()
                    tribal106cb.deselect()
                    #select PJD "attachment A" template 
                    pjdtscb.select()
                if data.text == 'Off':
                    pass
                           
                  
    # loads users setting window
    def loadSettings(self):
        USER.grid()
       
    # check box rule checker function
    def cb(self, CB):
        #This checks to make sure only one JD option is selected
        if CB == 'JD' and reqtypjd.get() == True:
            JD.grid()
        elif CB == 'JD' and reqtypjd.get() == False:
            JD.grid_remove()
            AJD.grid_remove()
            #deselect AJD
            typeajdcb.deselect()
            #deselect PJD
            typepjdcb.deselect()
            #reset uplands only
            uppresentcb.deselect()
            #reset wetland present
            wetpresentcb.deselect()
            #reset valid JD
            # validjdcb.deselect()
        if CB == 'GP' and reqtypgp.get() == True:
            GP.grid()
        elif CB == 'GP' and reqtypgp.get() == False:
            GP.grid_remove()
        #if approved JD then show the AJD screen/uncheck PJD box
        if CB == 'AJD' and typeajd.get() == True:
            #show AJD screen
            #AJD.grid()
            #make sure PJD is not checked
            typepjdcb.deselect()
        elif CB == 'AJD' and typeajd.get() == False:
            #remove AJD grid from view
            AJD.grid_remove()
        #if prliminary JD then uncheck PJD box/hide AJD grid
        if CB == 'PJD' and typepjd.get() == True:
            #reset uplands only
            uppresentcb.deselect()
            #make sure AJD is not checked
            typeajdcb.deselect()
            #hide AJD grid
            AJD.grid_remove()
            waterstypTNWcb.deselect() 
            waterstypTNWWcb.deselect() 
            waterstypTNWNRPWcb.deselect() 
            waterstypeRPWcb.deselect() 
            waterstypeRPWWDcb.deselect() 
            waterstypeRPWWNcb.deselect() 
            waterstypNRPWWNcb.deselect() 
            waterstypeIMPDcb.deselect() 
            waterstypeISOLcb.deselect()
            waterstypeNONJDcb.deselect()
        #if only uplands then select AJD and unselect PJD
        if CB == 'UPP' and uppresent.get() == True:
            #make sure PJD is not checked
            typepjdcb.deselect()
            #select AJD
            typeajdcb.select()
            #show AJD screen
            #AJD.grid()
            #deselect wetlands present
            wetpresentcb.deselect()
            #deselect jursidcitional authority
            regauth10.deselect()
            regauth404.deselect()
            waterstypTNWcb.deselect() 
            waterstypTNWWcb.deselect() 
            waterstypTNWNRPWcb.deselect() 
            waterstypeRPWcb.deselect() 
            waterstypeRPWWDcb.deselect() 
            waterstypeRPWWNcb.deselect() 
            waterstypNRPWWNcb.deselect() 
            waterstypeIMPDcb.deselect() 
            waterstypeISOLcb.deselect() 
        elif CB == 'UPP' and uppresent.get() == False:
            #reset uplands only
            uppresentcb.deselect()
            #deselect AJD
            typeajdcb.deselect()
            #show AJD screen
            AJD.grid_remove()         
         #if wetlands present
        if CB == 'WETP' and wetpresent.get() == True:
            #reset uplands only
            uppresentcb.deselect() 
        if CB == 'WETP' and wetpresent.get() == False: 
            waterstypTNWWcb.deselect() 
            waterstypeRPWWDcb.deselect() 
            waterstypeRPWWNcb.deselect() 
            waterstypNRPWWNcb.deselect() 
            waterstypeISOLcb.deselect() 
        # if ESA then open the ESA window
        if CB == 'ESAEFHMSA' and esaefhmsa.get() == True:
            #show AJD screen
            ESA.grid()
            #make sure PJD is not checked
            # typepjdcb.deselect()
        elif CB == 'ESAEFHMSA' and esaefhmsa.get() == False:
            #remove AJD grid from view
            ESA.grid_remove()
        # if 160 then open 106 window
        if CB == 'TRIB106' and tribal106.get() == True:
            #show AJD screen
            CR.grid()
            #make sure PJD is not checked
            # typepjdcb.deselect()
        elif CB == 'TRIB106' and tribal106.get() == False:
            #remove AJD grid from view
            CR.grid_remove()
            
        # if GP and JD 
        if CB == 'GP' or CB == 'JD':
            #deselect NPR, NOV, Non-compliance, and Scoping
            reqtypenprcb.deselect()
            # reqtypenovcb.deselect()
            reqtypenoncompcb.deselect()
            # reqtypescpcb.deselect()        
            # if both are true select both gp and gp jd
            if reqtypgp.get() == True and reqtypjd.get() == True:
                # select needed templates
                gptscb.select()
                gpjdtscb.select()
                cpfurntscb.select()
                appltscb.select()
                ddoctscb.select()
                cftscb.select()
                #deselect unneeded templates
                jdonlytscb.deselect()
                # scopingtscb.deselect()
                nprtscb.deselect()
                noncomptscb.deselect()
                # novtscb.deselect()
             # if GP false and JD false
            elif reqtypgp.get() == False and reqtypjd.get() == False:
                gptscb.deselect()
                gpjdtscb.deselect()
                cpfurntscb.deselect()
                sctscb.deselect()
                cftscb.deselect()
                appltscb.deselect()
                ddoctscb.deselect()
                mtftscb.deselect()
                raptscb.deselect()
                pjdtscb.deselect()
                jdonlytscb.deselect()
                noncomptscb.deselect()
                nprtscb.deselect()
                # scopingtscb.deselect()
                #deselect all gp options
                compmitcb.deselect()
                compmitprmcb.deselect()
                specialcondcb.deselect()
                atfcb.deselect()
                rwaivercb.deselect()
                waivercb.deselect()
                imp408cb.deselect()
                impwscb.deselect()
                agencycordcb.deselect()
                corpscordcb.deselect()
                esaefhmsacb.deselect()
                tribal106cb.deselect()
            # if GP is unselected but JD is still present, select JD only
            elif reqtypgp.get() == False and reqtypjd.get() == True:
                # select needed tamples
                cpfurntscb.select()
                appltscb.select()
                jdonlytscb.select()
                #deselect unneeded templates
                gptscb.deselect()
                gpjdtscb.deselect()
                ddoctscb.deselect()
                cftscb.deselect()
                # scopingtscb.deselect()
                nprtscb.deselect()
                noncomptscb.deselect()
                #deselect GP otions
                compmitcb.deselect()
                compmitprmcb.deselect()
                specialcondcb.deselect()
                atfcb.deselect()
                rwaivercb.deselect()
                waivercb.deselect()
                imp408cb.deselect()
                impwscb.deselect()
                agencycordcb.deselect()
                corpscordcb.deselect()
                esaefhmsacb.deselect()
                tribal106cb.deselect()
                # if GP only true
            elif reqtypgp.get() == True:
                gptscb.select()
                cpfurntscb.select()
                appltscb.deselect()
                ddoctscb.select()
                cftscb.select()
                #deselect unneeded templates
                gpjdtscb.deselect()
                jdonlytscb.deselect()
                # scopingtscb.deselect()
                nprtscb.deselect()
                noncomptscb.deselect()
                # novtscb.deselect()
        #select for NOV - can include ATF JD and GP
        if CB == 'NOV' and reqtypnov.get() == True:
            # deselect incompatiable action types
            reqtypenprcb.deselect()
            reqtypenoncompcb.deselect()
            # reqtypescpcb.deselect() 
            #deselect incompatiable templates
            noncomptscb.deselect()
            nprtscb.deselect()
            # scopingtscb.deselect()
            #select NOV 
            novtscb.select()
        elif CB == 'NOV' and reqtypnov.get() == False:
            #deselect NOV 
            novtscb.deselect()
        # if NPR
        if CB == 'NPR' and reqtypnpr.get() ==True:
            #deselect JD, GP, NOV, Non-compliance, and Scoping
            reqtypegpcb.deselect()
            GP.grid_remove()
            reqtypejdcb.deselect()
            JD.grid_remove()
            reqtypenovcb.deselect()
            reqtypenoncompcb.deselect()
            # reqtypescpcb.deselect()   
            #deselect all templates except npr
            gptscb.deselect()
            gpjdtscb.deselect()
            cpfurntscb.deselect()
            sctscb.deselect()
            cftscb.deselect()
            appltscb.deselect()
            ddoctscb.deselect()
            mtftscb.deselect()
            raptscb.deselect()
            pjdtscb.deselect()
            jdonlytscb.deselect()
            novtscb.deselect()
            noncomptscb.deselect()
            # scopingtscb.deselect()
            #select NPR 
            nprtscb.select()
        elif CB == 'NPR' and reqtypnpr.get() ==False:
            #deselect JD, GP, NOV, Non-compliance, and Scoping
            reqtypegpcb.deselect()
            GP.grid_remove()
            reqtypejdcb.deselect()
            JD.grid_remove()
            reqtypenovcb.deselect()
            reqtypenoncompcb.deselect()
            # reqtypescpcb.deselect()   
            #deselect all templates except npr
            gptscb.deselect()
            gpjdtscb.deselect()
            cpfurntscb.deselect()
            sctscb.deselect()
            cftscb.deselect()
            appltscb.deselect()
            ddoctscb.deselect()
            mtftscb.deselect()
            raptscb.deselect()
            pjdtscb.deselect()
            jdonlytscb.deselect()
            novtscb.deselect()
            noncomptscb.deselect()
            # scopingtscb.deselect()
            #deselect NPR 
            nprtscb.deselect()
        # if non-compliance
        if CB == 'NC' and reqtypnoncomp.get() ==True:
            #deselect JD, GP, NOV, Non-compliance, and Scoping
            reqtypegpcb.deselect()
            GP.grid_remove()
            reqtypejdcb.deselect()
            JD.grid_remove()
            reqtypenovcb.deselect()
            reqtypenprcb.deselect()
            # reqtypescpcb.deselect()   
            #deselect all templates except npr
            gptscb.deselect()
            gpjdtscb.deselect()
            cpfurntscb.deselect()
            sctscb.deselect()
            cftscb.deselect()
            appltscb.deselect()
            ddoctscb.deselect()
            mtftscb.deselect()
            raptscb.deselect()
            pjdtscb.deselect()
            jdonlytscb.deselect()
            novtscb.deselect()
            nprtscb.deselect()
            # scopingtscb.deselect()
            #select NPR 
            noncomptscb.select()
        elif CB == 'NC' and reqtypnoncomp.get() ==False:
            #deselect JD, GP, NOV, Non-compliance, and Scoping
            reqtypegpcb.deselect()
            GP.grid_remove()
            reqtypejdcb.deselect()
            JD.grid_remove()
            reqtypenovcb.deselect()
            reqtypenprcb.deselect()
            # reqtypescpcb.deselect()   
            #deselect all templates except npr
            gptscb.deselect()
            gpjdtscb.deselect()
            cpfurntscb.deselect()
            sctscb.deselect()
            cftscb.deselect()
            appltscb.deselect()
            ddoctscb.deselect()
            mtftscb.deselect()
            raptscb.deselect()
            pjdtscb.deselect()
            jdonlytscb.deselect()
            novtscb.deselect()
            nprtscb.deselect()
            # scopingtscb.deselect()
            #deselect NPR 
            noncomptscb.deselect()
        # if scoping
        # if CB == 'SCP' and reqtypscp.get() ==True:
        #     #deselect JD, GP, NOV, Non-compliance, and Scoping
        #     reqtypegpcb.deselect()
        #     GP.grid_remove()
        #     reqtypejdcb.deselect()
        #     JD.grid_remove()
        #     reqtypenovcb.deselect()
        #     reqtypenprcb.deselect()
        #     reqtypenoncompcb.deselect()   
        #     #deselect all templates except npr
        #     gptscb.deselect()
        #     gpjdtscb.deselect()
        #     cpfurntscb.deselect()
        #     sctscb.deselect()
        #     cftscb.deselect()
        #     appltscb.deselect()
        #     ddoctscb.deselect()
        #     mtftscb.deselect()
        #     raptscb.deselect()
        #     pjdtscb.deselect()
        #     jdonlytscb.deselect()
        #     novtscb.deselect()
        #     nprtscb.deselect()
        #     noncomptscb.deselect()
        #     #select NPR 
        #     scopingtscb.select()
        # elif CB == 'SCP' and reqtypscp.get() ==False:
        #     #deselect JD, GP, NOV, Non-compliance, and Scoping
        #     reqtypegpcb.deselect()
        #     GP.grid_remove()
        #     reqtypejdcb.deselect()
        #     JD.grid_remove()
        #     reqtypenovcb.deselect()
        #     reqtypenprcb.deselect()
        #     reqtypenoncompcb.deselect()  
        #     #deselect all templates except npr
        #     gptscb.deselect()
        #     gpjdtscb.deselect()
        #     cpfurntscb.deselect()
        #     sctscb.deselect()
        #     cftscb.deselect()
        #     appltscb.deselect()
        #     ddoctscb.deselect()
        #     mtftscb.deselect()
        #     raptscb.deselect()
        #     pjdtscb.deselect()
        #     jdonlytscb.deselect()
        #     novtscb.deselect()
        #     nprtscb.deselect()
        #     noncomptscb.deselect()
        #     #deselect NPR 
        #     scopingtscb.deselect()
        #if compensatory mitigation then select compensatory mitigation checkbox, special conditions, special condition template, and mtf template 
        if CB == 'COMPMIT' and compmit.get() ==True:
            mtftscb.select()
            specialcondcb.select()
            sctscb.select()
        if CB == 'COMPMIT' and compmit.get() ==False:
            # specialcondcb.deselect()
            # sctscb.deselect()
            compmitprmcb.deselect()
            mtftscb.deselect()
        #if SC then select SC - if unslecect then unselect compensatory mitigation
        if CB == 'SPECON' and specialcond.get() ==True:
            specialcondcb.select()     
            sctscb.select()                
        if CB == 'SPECON' and specialcond.get() ==False:
            specialcondcb.deselect()
            sctscb.deselect()
            compmitprmcb.deselect()
            mtftscb.deselect()
            compmitcb.deselect()
        if CB == 'IMPWS' or CB == 'WAIVER' or CB == 'TRIB106' or CB == 'ESAEFHMSA':
            # if affect to wild and secnice then agency coordination
            if impws.get() ==True:
                agencycordcb.select()             
            if impws.get() == False:
                if waiver.get() == False and esaefhmsa.get()==False and tribal106.get()==False:
                    agencycordcb.deselect()
            # if waiver then agency coordination
            if waiver.get() ==True:
                agencycordcb.select()            
            if waiver.get() ==False:
                if impws.get() ==False and esaefhmsa.get()==False and tribal106.get()==False:
                    agencycordcb.deselect()
            # Affect to ESA select agency coordination
            if esaefhmsa.get() ==True:
                agencycordcb.select()            
            if esaefhmsa.get() ==False:
                if waiver.get()==False and impws.get() ==False and tribal106.get()==False:
                    agencycordcb.deselect()
            # affect to 106 then select agency coordination
            if tribal106.get() ==True:
                agencycordcb.select()            
            if tribal106.get() ==False:
                if waiver.get()==False and impws.get() ==False and esaefhmsa.get()==False:
                    agencycordcb.deselect()
        # affect to 408 then select internal corps
        if CB == 'IMP408' and imp408.get() ==True:
            corpscordcb.select()                          
        if CB == 'IMP408' and imp408.get() ==False:
            corpscordcb.deselect()
        # if AJD select rapanos template
        if CB == 'AJD' and typeajd.get() ==True:
            raptscb.select()   
            pjdtscb.deselect()                    
        if CB == 'AJD' and typeajd.get() ==False:
            raptscb.deselect()
        # if PJD select PJD form
        if CB == 'PJD' and typepjd.get() ==True:
            pjdtscb.select()   
            raptscb.deselect()                    
        if CB == 'PJD' and typepjd.get() ==False:
            pjdtscb.deselect()
        #  if section 404 or 10 then define authority
        if CB == '404' and authtyp404.get() ==True:
            authority.set(True);                    
        if CB == '404' and authtyp404.get() ==False:
            if authtyp10.get()==False:
                authority.set(False);  
        if CB == '10' and authtyp10.get() ==True:
            authority.set(True);                   
        if CB == '10' and authtyp10.get() ==False:
            if authtyp404.get()==False:
               authority.set(False); 
        # if any of the rapanos wetlands are check then check wetpresent
        if CB == 'RWETLAND' and wetpresent.get() ==False and uppresent.get() == False:
            wetpresentcb.select()
        # if a rapanos water is checked when uplands only is selected then unslect all rapanos waters
        if (CB == 'RWETLAND' or CB == 'RWATER') and uppresent.get() == True:
            waterstypTNWcb.deselect() 
            waterstypTNWWcb.deselect() 
            waterstypTNWNRPWcb.deselect() 
            waterstypeRPWcb.deselect() 
            waterstypeRPWWDcb.deselect() 
            waterstypeRPWWNcb.deselect() 
            waterstypNRPWWNcb.deselect() 
            waterstypeIMPDcb.deselect() 
            waterstypeISOLcb.deselect() 
    
if __name__ == "__main__":
    root = tk.Tk()
    MainApplication(root).grid(column=0, row=0)
    #prevent user resizing
    root.resizable(False, False)
    try:
        root.mainloop()
    except SystemExit as e:
        print ('Error!' + e)
        print ('Press enter to exit (and fix the problem)')
        raw_input()
