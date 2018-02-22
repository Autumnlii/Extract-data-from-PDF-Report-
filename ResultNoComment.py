#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
    Created on Thu Feb 9 18:35:37 2018
    
    @author: Qiuying(Autumn) Li
    
    """
import sys
import importlib
import re
importlib.reload(sys)

from pdfminer.pdfparser import PDFParser,PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal,LAParams
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed
from openpyxl import Workbook

from openpyxl import load_workbook

from openpyxl.writer.excel import ExcelWriter 

#print ("--------------")
def write_xlsx(content=[]):
    init_keys = "Course	StudentName	Grade	ReportDate	LASID	SchoolCode	SchoolName	DOB	DistrictCode	DistrictName	ScaleScore	AchievementLevel	RescoreFlag	Lower	Upper	AchievementLevel(bottom)".split()
    used_keys = ["name",
                    "LASID",
                    "DOB",
                    "Grade",
                    "RD",
                    "School",
                    "District",
                    "Score",
                    "Score_level",
                    "low_top",
                    "course"]
           
    wb = Workbook()
    ws = wb.active
    ws.append(init_keys)
    init_keys = "Course	StudentName	Grade	ReportDate	LASID	SchoolCode	SchoolName	DOB	DistrictCode	DistrictName	ScaleScore	AchievementLevel	RescoreFlag	Lower	Upper	AchievementLevel(bottom)".split()
    for item in content:
        course = item.get("course"," ")
        LASID = item.get("LASID"," ")
        Grade = item.get("Grade"," ")
        RD = item.get("RD"," ")
        School = item.get("School"," ")
        District = item.get("District"," ")
        Score = item.get("Score"," ") #invalid
        Score_level = item.get("Score_level"," ")
        low_top = item.get("low_top"," ")
        name = item.get("name"," ")
        DOB = item.get("DOB"," ")
        level = " "
        temp = []
        temp.append(course)
        temp.append(name)
        temp.append(Grade)
        temp.append(RD)
        temp.append(LASID)
        if School:
            School_list = School.split()
            temp.append(School_list[0])
            temp.append(" ".join(School_list[1:]))
        else:
            temp.append(" ");temp.append(" ")
        temp.append(DOB)
        if District:
            District_list = District.split()
            temp.append(District_list[0])
            temp.append(" ".join(District_list[1:]))
        else:
            temp.append(" ");temp.append(" ")
        if Score_level:
            temp.append(Score_level[0])
            temp.append(Score_level[1])
            level = Score_level[1]
        else:
            temp.append(" ");temp.append(" ") 
        temp.append(" ")
        if low_top:
            temp.append(low_top[0])
            temp.append(low_top[1])
        else:
            temp.append(" ");temp.append(" ") 
        
        temp.append(level)
        ws.append(temp)
        #temp.append(course)
        
        
        
    wb.save("result.xlsx") 

def get_name(name_str):
    name_str_list = name_str.split("\n")
    name_p = re.compile("^Student:\s+(.*)")
    for name in name_str_list:
        name_s = name_p.search(name)
        if name_s:
            return name_s.group(1)
            
def get_LASID(LASID_str):
    LASID_str_list = LASID_str.split("\n")
    LASID_p = re.compile("^LASID:\s+(.*)")
    for LASID in LASID_str_list:
        LASID_s = LASID_p.search(LASID)
        if LASID_s:
            return LASID_s.group(1)
def get_DOB(n_str):
    name_str_list = n_str.split("\n")
    name_p = re.compile("^Date of Birth:\s+(.*)")
    for name in name_str_list:
        name_s = name_p.search(name)
        if name_s:
            return name_s.group(1)
def get_Grade(n_str):
    name_str_list = n_str.split("\n")
    name_p = re.compile("^Grade:\s+(.*)")
    for name in name_str_list:
        name_s = name_p.search(name)
        if name_s:
            return name_s.group(1)
  
def get_RD(n_str):
    name_str_list = n_str.split("\n")
    name_p = re.compile("^Report Date:\s+(.*)")
    for name in name_str_list:
        name_s = name_p.search(name)
        if name_s:
            return name_s.group(1)
def get_School(n_str):
    name_str_list = n_str.split("\n")
    name_p = re.compile("^School:\s+(.*)")
    for name in name_str_list:
        name_s = name_p.search(name)
        if name_s:
            return name_s.group(1)
def get_District(n_str):
    name_str_list = n_str.split("\n")
    name_p = re.compile("^District:\s+(.*)")
    for name in name_str_list:
        name_s = name_p.search(name)
        if name_s:
            return name_s.group(1)
def get_Score(n_str):
    name_str_list = n_str.split("\n")
    name_p = re.compile("The student's score is\s+(\d+)")
    for name in name_str_list:
        name_s = name_p.search(name)
        if name_s:
            return name_s.group(1)      

def get_Score_level(n_str):
    name_str_list = n_str.split("\n")
    name_p = re.compile("The student.*?score is\s+(\d+),\s+which falls in the (.*?) achievement level")
    for name in name_str_list:
        name_s = name_p.search(name)
        if name_s:
            return name_s.group(1),name_s.group(2)   
  
def get_low_top(n_str):  
    name_str_list = n_str.split("\n")
    name_p = re.compile("the score would fall in the range of (\d+) to (\d+)")
    for name in name_str_list:
        name_s = name_p.search(name)
        if name_s:
            return name_s.group(1),name_s.group(2)
            
path = 'a.pdf'



def parse(file_name):
    fp = open(file_name, 'rb')

    praser = PDFParser(fp)

    doc = PDFDocument()

    praser.set_document(doc)
    doc.set_parser(praser)
    useful = []
 
    doc.initialize()

    if not doc.is_extractable:
        raise PDFTextExtractionNotAllowed
    else:
        
        rsrcmgr = PDFResourceManager()
     
        laparams = LAParams()
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        
        interpreter = PDFPageInterpreter(rsrcmgr, device)


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
        for page in doc.get_pages():
            interpreter.process_page(page)
            
            layout = device.get_result()
            
            read_flag = 0
            for x in layout:
                if (isinstance(x, LTTextBoxHorizontal)):
                    results = x.get_text()
                    print (results)
                    if page_number%2 == 0 and read_flag==0:
                        #temp_use.append(results)
                        temp_dict["course"] = results.split("\n")[0]
                        read_flag = 1
                        continue
                    else:
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
                            #print ("hhhh")
                        if get_low_top(results):
                            temp_dict["low_top"] = get_low_top(results)
                    #print (temp_dict)
                    #input("==")
            #page_number += 1
            #print (page_number)
            #if page_number%2 == 0:
            if 1:
                #print (temp_dict)
                useful.append(temp_dict)
                #input("=======")
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
    for f in ['a.pdf','b.pdf']:
        used += parse(f)
    write_xlsx(used)
