######################################
##  ------------------------------- ##
##           WordXML Class          ##
##  ------------------------------- ##
##     Written by: David Shaeffer   ##
##  ------------------------------- ##
##      Last Edited on: 12-18-2019   ##
##  ------------------------------- ##
######################################
"""This is the WordXML Class."""

from lxml import etree as ET
from copy import deepcopy
from tkinter import filedialog
import os


class WordXML:
    """This is the class doc string."""

    def __init__(self, xmlfile):
        """This takes the xml file path passed to the class and parses the xml using the lxml module."""
        self.xmlfile = ET.parse(xmlfile).getroot()

    def savefile(self):
        """This function saves the xml to a new file in a user defined location."""
        try:
            """Create a new xml element tree and write the updated xml to the new tree instance."""
            tree = ET.ElementTree(self.xmlfile)
            """
      Open a file dialog box using tkinter and set the default extension to xml.
      """
            f = filedialog.asksaveasfilename(defaultextension=".xml",title = "Save Word Template",filetypes = (("xml files","*.xml"),("all files","*.*")))
            """Write the new xml tree instance to the xml file"""
            tree.write(f, xml_declaration=True, encoding="UTF-8",
                       method="xml", standalone="yes")
            """Use the OS to open the xml file"""
            os.startfile(f)
        except:
            """If the user presses cancels or it cannot save for some reason then pass."""
            print("Could not save template.")

    def editcdp(self, fieldname, fieldtext):
        """This functions modifies the custom document properties."""
        try:
            """Parse through the xml document and find the custom document property by tag."""
            for CDP in self.xmlfile.iter('{CustomDocumentProperties}' + fieldname):
                """Insert the text (field name) passed to the function to the custom document property specified (field name)."""
                CDP.text = fieldtext
                CDP.set('updated', 'yes')
        except:
            print("Could not modify custome document property.")
            raise

    def removesection(self, x):
        """This function assigns sections to all paragraphs before a page break and deletes the section
        specified by the x function parameter"""
        try:
            i = 0
            for body in self.xmlfile.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body'):
                """Iterate through all the paragraphs in the document body."""
                for child in body:
                    """If sections already exist, break """
                    if child.get('section'):
                        break
                    else:
                        """Create a new temporary section property on each paragraph and sets it equal to i."""
                        child.set('section', str(i))
                        """Iterate through the document to determine the number of section breaks (scetPr)."""
                        for sections in child.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr'):
                            """For each section break increment i."""
                            i += 1
        except:
            raise
        try:
            for body in self.xmlfile.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body'):
                """Iterate through all the paragraphs in the document body."""
                for child in body:
                    """If the section value equals the section passed to the function (x) then remove all paragraphs in that section."""
                    if int(child.get('section')) == x:
                        child.getparent().remove(child)
        except:
            raise

    def removecdp(self, fieldname):
        """This functions removes the custom document property based on the properties name."""
        try:
            """Parse through the xml document and find all the data bindings."""
            for CDP in self.xmlfile.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}dataBinding'):
                """Search the xpath to determine if the databindings is the targetted custome document property."""
                if fieldname in CDP.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}xpath'):
                    """Delete the custom document property up to the sdt element"""
                    CDP.getparent().getparent().getparent().remove(CDP.getparent().getparent())
        except:
            print("Could not modify custom document property.")
            raise
    
    def removeCC(self, controlname):
        for sdt in self.xmlfile.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt'):
            for child in sdt:
                # find the correct control
                for tag in child.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tag'):
                    if tag.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') == controlname:
                        # if control is correct give the content control a name at the parents parent
                        tag.getparent().getparent().set('cc', str(controlname))
                        # go back through the root to find the control that matches the desired control name you just assined then delete it
                        for control in self.xmlfile.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt'):
                            if control.get('cc') == str(controlname):
                                control.getparent().remove(control)

    """This function takes a standrd content control and a list and inserts the data into the table"""
    def addtotable(self, SCC, data):
        numrows = len(data)-1
        try:
            """Find an existing table within the defined content control and add empty rows"""
            for sdt in self.xmlfile.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt'):
                for child in sdt:
                    # find the correct control
                    for tag in child.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tag'):
                        if tag.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') == SCC:
                            """Set a counter to determine which row you're on when irrirating"""
                            row = 0
                            for element in tag.getparent().getparent().iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr'):
                                row += 1
                                """The first row is the heading and the second is a blank row - if it is the second row we need to copy that for each record in the list"""
                                if row == 2:
                                    records = 0
                                    for record in data:
                                        records += 1
                                        for table in tag.getparent().getparent().iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl'):
                                            """Copy the row"""
                                            if records <= numrows:
                                                table.append(deepcopy(element))

            """Copy the data from the function argument into the blank rows"""
            for sdt in self.xmlfile.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt'):
                for child in sdt:
                    # find the correct control
                    for tag in child.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tag'):
                        if tag.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') == SCC:
                            """ Count Rows """
                            i = 0
                            for rows in tag.getparent().getparent().iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr'):
                                """Skip the heading row"""
                                if i > 0:
                                    if i <= len(data):
                                        """Count the number of columns"""
                                        a = 0
                                        for text in rows.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                                            """ Assgin the correct data from the JD form to the correct table columns"""
                                            if a == 0:
                                                text.text = data[i-1][0]
                                            if a == 1:
                                                text.text = data[i-1][1]
                                            if a == 2:
                                                text.text = data[i-1][2]
                                            if (a == 3):
                                                text.text = data[i-1][3] + \
                                                    " " + data[i-1][4]
                                            if (a == 4):
                                                """If the cowardin class is wetland then specify wetland else specify non-wetland"""
                                                if "EM" in data[i-1][6] or "FO" in data[i-1][6] or "R6" in data[i-1][6] or "RP" in data[i-1][6] or "SS" in data[i-1][6]:
                                                    text.text = "Wetland"
                                                else:
                                                    text.text = "Non-Wetland"
                                            if (a == 5):
                                                text.text = data[i-1][5]
                                            a += 1
                                i += 1
        except:
            print("Could not modify table.")
            raise



    """ This function changes gray highlight to yellow based on the index"""
    def addHL(self, index):
        i = 0
        # Find all ffData
        for hi in self.xmlfile.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight'):
            if hi.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') == "lightGray" or hi.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') == "yellow":
                if i == index:
                    hi.set(
                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'yellow')
                # else:
                #     hi.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'none')
                i += 1

    """ This function changes the dropdown by index """

    def updatedropdown(self, ddTag, value):
        i = 0
        x = 0
        global item
        for sdt in self.xmlfile.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt'):
            for child in sdt:
                # find the correct tag
                for tag in child.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tag'):
                    if tag.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') == ddTag:
                        # if tag is correct, set to index value based on use input
                        #### FOR DROP DOWN LISTS ###
                        for dropdown in child.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}dropDownList'):
                            # print dropdown
                            for listItem in dropdown:
                                # print listItem
                                i += 1
                                if i == value:
                                    # set last value item based on specified item value
                                    item = listItem.get(
                                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}displayText')
                                    dropdown.set(
                                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lastValue', item)
                                    # set the content to the selected item
                                    for sdtContent in sdt.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtContent'):
                                        for content in sdtContent.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                                            content.text = item
                                            # print item
                                        for it in sdtContent.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}i'):
                                            it.getparent().remove(it)
                                        for color in sdtContent.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color'):
                                            color.set(
                                                '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '000000')
                                    # updated custom document property
                                    self.editcdp(ddTag, item)
                        #### FOR COMBO BOXES ###
                        for combobox in child.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comboBox'):
                            for listItem in combobox:
                                i += 1
                                if i == value:
                                    # set last value item based on specified item value
                                    item = listItem.get(
                                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}displayText')
                                    # print item
                                    combobox.set(
                                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lastValue', item)
                                    # set the content to the selected item
                                    for sdtContent in sdt.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtContent'):
                                        for content in sdtContent.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                                            # TODO fix waiver No.No. issue
                                            content.text = item
                                        for it in sdtContent.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}i'):
                                            it.getparent().remove(it)
                                        for color in sdtContent.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color'):
                                            color.set(
                                                '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '000000')
                                    # updated custom document property
                                    self.editcdp(ddTag, item)
