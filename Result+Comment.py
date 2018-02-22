#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
    Created on Thu Feb 9 18:35:37 2018
    
    @author: Qiuying(Autumn) Li
    
    """
import sys
import importlib
import re
import lib.reload(sys)

""" Summary of the all 12 methods:
    Below code will scan PDF file and store data into excel file.
    # Make sure your python version is 3.5 
    1. use pdfminer to read pdf file (https://github.com/pdfminer/pdfminer.six)
    2. use openpyxl to write and generate xlsx file in the python (https://pypi.python.org/pypi/openpyxl)
    3. There are 12 methods in this file:
    method # 1. def write_xlsx(content=[]) is to write the excel file;
    method # 2. def get_name(name_str): is to detect and store student name in the file;
    methods #3 - method #11 is to get useful information of students(LASID, DOB, Grade, RD, School, District, Score, Score level)

    method #12. def parse(file_name): to exact text information from PDF, and call method above to detect all the useful information of the students.
    
    Test tips:
    1. make sure you installed pdfminer and openpyxl;
    2. make sure that all pdf files are in the same path;
    3. After running the code, excel file: result.xlsx will be generated.
    
    """

from pdfminer.pdfparser import PDFParser,PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal,LAParams
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.writer.excel import ExcelWriter 


def write_xlsx(content=[]):
    """ write_xlsx can write data for each variables we need for reports
        
        Method description:
        1.Initially define the formal varibales names will be used in excel export;
        2.Define the short variable names in the code for the reaons of convenience;
        3.Create a workbook and write data into the work book
        4.Use for loop to store data corresponding to variables one by one, with specific order
        5.Some variables can be read at once(course, LASID,Grade,RD...etc)
        6.Some varibales have more than a value, such as "School" contains(SchoolCode and ShoolName)
        7.Save the workbook as result.xlsx
        
        return: excel file with name: "result.xlsx"
        
        raises: none know bugs
        
        """
    init_keys = "Course	StudentName	Grade	ReportDate	LASID	SchoolCode	SchoolName	DOB	DistrictCode	DistrictName	ScaleScore	AchievementLevel	RescoreFlag	Lower	Upper	AchievementLevel(bottom)".split()
    used_keys = ["name",
                    "LASID",
                    "DOB",
                    "Grade",
                    "RD",
                    "School", # contains info of SchoolCode and SchoolName
                   "District", #Contains info of DistrictCode and DistrictName
                    "Score",
                   "Score_level", # contains infor of AchievementLevel and AchievementLevel(bottom)
                    "low_top",#contains info of Lower bound and Upper bound
                    "course"]
           
    #create a new workbook name as wb
    wb = Workbook()
    # grab the active worksheet
    ws = wb.active
    #Rows of initial variable names can be appended
    ws.append(init_keys)
    init_keys = "Course	StudentName	Grade	ReportDate	LASID	SchoolCode	SchoolName	DOB	DistrictCode	DistrictName	ScaleScore	AchievementLevel	RescoreFlag	Lower	Upper	AchievementLevel(bottom)".split()

    for item in content:
        course = item.get("course"," ") # get course title
        LASID = item.get("LASID"," ")
        Grade = item.get("Grade"," ")
        RD = item.get("RD"," ")
        School = item.get("School"," ")
        District = item.get("District"," ")
        Score = item.get("Score"," ")
        Score_level = item.get("Score_level"," ")
        low_top = item.get("low_top"," ")
        name = item.get("name"," ")
        DOB = item.get("DOB"," ")
        level = " " # it belongs to part of score_level variable, no need to read twice
        temp = [] # create a new array to store student information with right order;
        temp.append(course) #store course title
        temp.append(name)
        temp.append(Grade)
        temp.append(RD)
        temp.append(LASID)
        
        #Below are the variables has more than a value, need a if statement to justify
        if School: #if the value of the school is not empty
            School_list = School.split()#split school_list
            temp.append(School_list[0])# store index[0] as SchoolCode
            temp.append(" ".join(School_list[1:]))#store index[1] as SchoolName
        else:
            temp.append(" ");temp.append(" ")#if school value is empty, skipd and write DOB
        temp.append(DOB)
        if District:
            District_list = District.split()
            temp.append(District_list[0]) #index[0] as DistrictCode
            temp.append(" ".join(District_list[1:]))#index[1] as DistrcitName
        else:
            temp.append(" ");temp.append(" ")
        if Score_level:
            temp.append(Score_level[0]) #index[0] as AchievementLevel
            temp.append(Score_level[1]) #index [1] as AchievementLevel(bottom)
            level = Score_level[1]
        else:
            temp.append(" ");temp.append(" ") 
        temp.append(" ")
        if low_top:
            temp.append(low_top[0]) #index[0] as lower
            temp.append(low_top[1]) #index[1] as upper
        else:
            temp.append(" ");temp.append(" ")

        temp.append(level)
        ws.append(temp)
        #temp.append(course)
        
        
        
    wb.save("result.xlsx")

def get_name(name_str):
    """ scan, detect and store data of student name;
        
        Method description:
        1. Split strings and sep = "/n";
        2. define the pattern of the student names, and find student name. Regular expression as reference
        3. if the splitted strings matches the student names we detected
        4. stored as student names
        
        return: student name
        
        raises: none know bugs
        
        """
    name_str_list = name_str.split("\n")
    name_p = re.compile("^Student:\s+(.*)")
    for name in name_str_list:
        name_s = name_p.search(name)
        if name_s:
            return name_s.group(1)
def get_LASID(LASID_str):
    """ scan, detect and store data of student LASID;
        
        Method description:
        1. Split strings and sep = "/n";
        2. define the pattern of the student LASID, and find student LASID. Regular expression as reference
        3. if the splitted strings matches the student LASID we detected
        4. stored as student LASID
        
        return: student LASID
        
        raises: none know bugs
        
        """
    LASID_str_list = LASID_str.split("\n")
    LASID_p = re.compile("^LASID:\s+(.*)")
    for LASID in LASID_str_list:
        LASID_s = LASID_p.search(LASID)
        if LASID_s:
            return LASID_s.group(1)
def get_DOB(n_str):
    """ scan, detect and store data of student DOB(DATE OF BIRTH);
        
        Method description:
        1. Split strings and sep = "/n";
        2. define the pattern of the student DOB, and find student LASID. Regular expression as reference
        3. if the splitted strings matches the student DOB we detected
        4. stored as student DOB
        
        return: student DOB
        
        raises: none know bugs
        
        """
    name_str_list = n_str.split("\n")
    name_p = re.compile("^Date of Birth:\s+(.*)")
    for name in name_str_list:
        name_s = name_p.search(name)
        if name_s:
            return name_s.group(1)
def get_Grade(n_str):
    """ scan, detect and store data of student Grade;
        
        Method description:
        1. Split strings and sep = "/n";
        2. define the pattern of the student Grade, and find student Grade. Regular expression as reference
        3. if the splitted strings matches the student Grade we detected
        4. stored as student Grade
        
        return: student Grade
        
        raises: none know bugs
        
        """
    name_str_list = n_str.split("\n")
    name_p = re.compile("^Grade:\s+(.*)")
    for name in name_str_list:
        name_s = name_p.search(name)
        if name_s:
            return name_s.group(1)
def get_RD(n_str):
    """ scan, detect and store data of RD(Report Date);
        
        Method description:
        1. Split strings and sep = "/n";
        2. define the pattern of the student RD, and find student RD. Regular expression as reference
        3. if the splitted strings matches the student RD we detected
        4. stored as student RD
        
        return: student RD
        
        raises: none know bugs
        
        """
    name_str_list = n_str.split("\n")
    name_p = re.compile("^Report Date:\s+(.*)")
    for name in name_str_list:
        name_s = name_p.search(name)
        if name_s:
            return name_s.group(1)
def get_School(n_str):
    """ scan, detect and store data of School information
        
        Method description:
        1. Split strings and sep = "/n";
        2. define the pattern of the student School, and find student School. Regular expression as reference
        3. if the splitted strings matches the student School we detected
        4. stored as student School
        
        return: student School
        
        raises: none know bugs
        
        """
    name_str_list = n_str.split("\n")
    name_p = re.compile("^School:\s+(.*)")
    for name in name_str_list:
        name_s = name_p.search(name)
        if name_s:
            return name_s.group(1)
def get_District(n_str):
    """ scan, detect and store data of District
        Method description:
        1. Split strings and sep = "/n";
        2. define the pattern of the student District, and find student District. Regular expression as reference
        3. if the splitted strings matches the student District we detected
        4. stored as student District
        
        return: student District
        
        raises: none know bugs
        
        """
    name_str_list = n_str.split("\n")
    name_p = re.compile("^District:\s+(.*)")
    for name in name_str_list:
        name_s = name_p.search(name)
        if name_s:
            return name_s.group(1)
def get_Score(n_str):
    """ scan, detect and store data of Score
        Method description:
        1. Split strings and sep = "/n";
        2. define the pattern of the student Score, and find student Score. Regular expression as reference
        3. if the splitted strings matches the student Score we detected
        4. stored as student Score
    
        return: student Score
        
        raises: none know bugs
        
        """
    name_str_list = n_str.split("\n")
    name_p = re.compile("The student's score is\s+(\d+)")
    for name in name_str_list:
        name_s = name_p.search(name)
        if name_s:
            return name_s.group(1)      

def get_Score_level(n_str):
    """ scan, detect and store data of Score level
        Method description:
        1. Split strings and sep = "/n";
        2. define the pattern of the student Score_level, and find student Score_level. Regular expression as reference
        3. if the splitted strings matches the student Score_level we detected
        4. stored as student Score_level
        
        return: student Score_level
        
        raises: none know bugs
        
        """
    name_str_list = n_str.split("\n")
    name_p = re.compile("The student.*?score is\s+(\d+),\s+which falls in the (.*?) achievement level")
    for name in name_str_list:
        name_s = name_p.search(name)
        if name_s:
            return name_s.group(1),name_s.group(2)   
  
def get_low_top(n_str):
    """ scan, detect and store data of Score lower bound and upper bound
        Method description:
        1. Split strings and sep = "/n";
        2. define the pattern of the student Score_low_top, and find student Score_low_top. Regular expression as reference
        3. if the splitted strings matches the student Score_low_top we detected
        4. stored as student Score_low_top
        
        return: student Score_low_top
        
        raises: none know bugs
        
        """
    name_str_list = n_str.split("\n")
    name_p = re.compile("the score would fall in the range of (\d+) to (\d+)")
    for name in name_str_list:
        name_s = name_p.search(name)
        if name_s:
            return name_s.group(1),name_s.group(2)
            
path = 'a.pdf'



def parse(file_name):
    """ Exact pdf file into text format and calling above methods to detect useful student information
        Method description:
        1. Exact pdf file and transform into text file by using pdfminer library functions
        2. Used above methods to detect and store the useful information in the list: uesful
        
        return: "useful", a list contains all the useful informaiton of student
        
        raises: none know bugs
        
        """
    
    fp = open(file_name, 'rb') #Open the file and read as binary mode;
    #Created pdf parser object to associate original pdf file
    praser = PDFParser(fp)
    # Created a blank PDF file to store useful information;
    doc = PDFDocument()
    #Connected parse object and doc.pdf
    praser.set_document(doc)
    doc.set_parser(praser)
    useful = [] # This is the empty list to store incoming useful student information
    #Initialize our empty doc
    doc.initialize()


    #To test if doc.pdf can be transformed as text format
    #If not, stopped;
    #Else, continue.
    if not doc.is_extractable:
        raise PDFTextExtractionNotAllowed
    else:
        #Created PDFRecourceManager to manage the shared resource
        rsrcmgr = PDFResourceManager()
        #Created PDF device object to store interpreted format of data
        laparams = LAParams()
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    
        #Create PDF interpreter object to transform shared informaiton in the rsrcmgr and stored in device ....
        interpreter = PDFPageInterpreter(rsrcmgr, device)

        #Use for loop to go through the file, and unit is page number
        #Initial page number is 0
        page_number = 0
        
        temp_use = []
        temp_dict = {
                    "name":"",
                    "LASID":"",
                    "DOB":"",
                    "Grade":"",
                    "RD":"",
                    "School":"",
                    "District":"",
                    "Score":"",
                    "Score_level":"",
                    "low_top":"",
                    "course":"",
                } 
        for page in doc.get_pages(): # doc.get_pages() to get page lists information
            interpreter.process_page(page)
            # To accept interpreted page LTPage object
            layout = device.get_result()
            #Layout means for every LTPage, which stores interpreted instance of corresponding pages, such asLTTextBox, LTFigure, LTImage, LTTextBoxHorizontal. If we need to capture strings, then it should be txt instance.
            read_flag = 0# this variable means the pages have been read
            for x in layout: #for every layout
                if (isinstance(x, LTTextBoxHorizontal)):# if the instance type of layout is LTTextBoxHorizontal
                    results = x.get_text()# we store all the txt in the results
                    print (results)
                    if page_number%2 == 0 and read_flag==0: # This is the even page number and no previous page has been read
                        #temp_use.append(results)
                        temp_dict["course"] = results.split("\n")[0]
                        read_flag = 1
                        continue
                    else:# if odd pages or there is previous page has been read
                        #  Continue getting name, LASID, DOB, etc...
                        if get_name(results):
                            temp_dict["name"] = get_name(results) 
                        if get_LASID(results):
                            temp_dict["LASID"] = get_LASID(results)
                        if get_DOB(results):
                            temp_dict["DOB"] = get_DOB(results)
                        if get_Grade(results):
                            temp_dict["Grade"] = get_Grade(results)
                        if get_RD(results):
                            temp_dict["RD"] = get_RD(results)
                        if get_School(results):
                            temp_dict["School"] = get_School(results)
                        if get_District(results):
                            temp_dict["District"] = get_District(results)
                        if get_Score(results):
                            temp_dict["Score"] = get_Score(results)
                        if get_Score_level(results):
                            temp_dict["Score_level"] = get_Score_level(results)
                        if get_low_top(results):
                            temp_dict["low_top"] = get_low_top(results)
                    #print (temp_dict)
                    #input("==")
            #page_number += 1
            #print (page_number)
            #if page_number%2 == 0:
            if 1:  # When page number == 1;
                useful.append(temp_dict)
                temp_dict = {
                    "name":"",
                    "LASID":"",
                    "DOB":"",
                    "Grade":"",
                    "RD":"",
                    "School":"",
                    "District":"",
                    "Score":"",
                    "Score_level":"",
                    "low_top":"",   
                } 
    return useful


if __name__ == '__main__':
    used = []
    for f in ['a.pdf','b.pdf']:# for two pdf files,
        used += parse(f)# conbine the information of two files;
    write_xlsx(used)       


""" Reference
    http://denis.papathanasiou.org/posts/2010.08.04.post.html
    https://pypi.python.org/pypi/pdfminer.six/20160202
    https://github.com/pdfminer/pdfminer.six
    https://pypi.python.org/pypi/openpyxl
    https://docs.python.org/2/library/re.html
    """
