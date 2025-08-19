# -*- coding: utf-8 -*-
"""
Created on Fri Aug  1 10:38:51 2025

@author: TAM4027
"""

import os
import re
import pandas as pd
import pdfplumber
import docx2txt
from win32com import client as wc
import doc2docx



def txt_from_pdf(file_path):
    '''
    converts pdfs to txt
    
    Args:
        file_path: (str) full filepath of attachment. Must be .pdf
    '''
    with pdfplumber.open(file_path) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


def doc_to_docx(dir_path):
    '''
    Converts all doc (old microsoft word document) to docx (more current)
    
    Args:
        dir_path : (str) filepath to directory.
    '''
    
    #ignores all "warning" files
    ignore_value = "warning"
    w = wc.Dispatch('Word.Application')
    w.Visible = False
    #try: --commented out. I want to know of errors
        
    #steps through directory and any subdirectories
    for root, dirs, files in os.walk(dir_path):
        #Thought: create sub folders for docx
#             dir_docx = root + r"/docx_converted"
#             if not os.path.exists(dir_docx):
#                 os.makedirs(dir_docx)
        
        # Iterate through any files found while stepping through directories
        for filename in files:
            file_path = os.path.join(root, filename)
            ext = os.path.splitext(filename)[1].lower()
            file_title = os.path.splitext(filename)[0]
            file_output = file_title + ".docx"
            output_path = os.path.join(root, file_output)
            
            #converts non-"Warning" .doc files that do not already exist
            if ext == ".doc" and ignore_value not in filename.lower() and not os.path.exists(output_path):
                doc2docx.convert(file_path, os.path.join(root, file_output))
                

                doc = w.Documents.Open(file_path)
                doc.SaveAs(output_path, 16)
                
                print(f"Successfully converted file {filename} to {file_output}")
                
                #Close doc and quit MS Word
                if doc:
                    doc.Close()
    if w:
        w.Quit()
                        
    # except Exception as e:
    #     print(f"Error processing {filename}: {e}")
        
    

#function to gather relevant info
##TODO review other items for item information

def parse_letter(text):
    
    ''' RegEx expressions to gather data and put in data dictionary
            Args:
                text (open text file): text file object to apply RegEx to
    '''
    #set up output dictionary
    result = {
        "recipient_name": None,
        "address": None,
        "dates": [],
        "general_laws": [],
        "license_actions" : [],
        "suspension_start_date" : None,
        "suspension_duration_days" : None}
    
        
    #preserve original for address and recipient
    text_og = text
    
    #replace all whitespace characters as " " for paragraph processing
    text_noNewLines = text.replace(r'\s+', ' ')
    
    #extract dates from text
    date_pattern = r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December) \d{1,2}, \d{4}"
    result["dates"] = re.findall(date_pattern, text_noNewLines)
    
    #create and clean lines
    lines = text_og.strip().splitlines()
    lines = [line.strip() for line in lines if line.strip()]
    
    #index of date line
    for i , line in enumerate(lines):
        if re.match(date_pattern, line):
            date_line_idx = i
            break
    
    #Recipient and address
    if date_line_idx is not None and date_line_idx + 1 < len(lines):
        result["recipient_name"] = lines[date_line_idx +1]
        
        address_lines = []
        for line in lines[date_line_idx +2]:
            if line == "":
                break #stop at empty line or section heading
            address_lines.append(line)
            if len(address_lines) >= 3:
                break #address should only be 2 lines long
        result["address"] = " ".join(address_lines)
        
    for i, line in enumerate(lines):
        if re.match(date_pattern, line):
            if i + 1 < len(lines):
                result["recipient_name"] = lines[i+1].strip() 
            if i + 3 < len(lines):
                result["address"] = lines[i+2].strip() + " " + lines[i +3].strip()
            break
    
    
    #G.L. chapters and sections
    #---------------- works but does not grab sections not directly after "G.L. c."
    # gl_pattern = r"G\.L\. c\. ?(\d+[A-Z]?),.*?(?:sections? )((?:\d+[A-Z]?,? ?)*(?:and )?\d+[A-Z]?)"
    # matches = re.findall(gl_pattern, text_noNewLines)
    # general_laws = []
    # for chapter, section_group in matches:
    #     section_group = section_group.replace(" and ", ",")
    #     sections = [s.strip() for s in section_group.split(",") if s.strip()]
        
    #     for section in sections:
    #         general_laws.append(f"G.L. c. {chapter}, section {section}")
    #-----------------
    
    #grab all sections (ignore G.L. c. text, just get sections)    
    
    gl_pattern = r"(?:sections?\s*)((?:\d+[A-Z]?(?:\s*\([^)]+\))*[\s,]*?)*(?:and\s+)?\d+[A-Z]?(?:\s*\([^)]+\))*)"
    gl_sections = re.findall(gl_pattern, text_noNewLines, flags=re.IGNORECASE | re.S)

    general_laws = []
        
    #normalize "and" to comma and split on comma. 
    for section in gl_sections:
        cleaned_section = section.replace("and", ",")
        split_sections = [s.strip() for s in cleaned_section.split(",") if s.strip()]
        
        for sec in split_sections:
            general_laws.append(f"section {sec}")
            
    result["general_laws"] = general_laws
    
    
    #license actions
    license_action_pattern = r"(suspending|revoking|reinstate .+) your (.+license)"
    license_actions = re.findall(license_action_pattern, text_noNewLines, re.IGNORECASE)
    result["license_actions"] = [
        "{} {}".format(action.lower(), license_type.lower())
        for action, license_type in license_actions
        ]
    
    
    #extract suspenstion date and duration
    suspension_date = re.search(
        r"effective(?: as of)?(?: 12:01 A\.M\.)?(?: on)? (?:[A-za-z]+, )?([A-Za-z]+ \d{1,2}, \d{4})",
        text_noNewLines, re.IGNORECASE
        )

    suspension_duration = re.search(
        r"for (?:a period of )?(\d+) days", 
        text_noNewLines, re.IGNORECASE
        )

    if suspension_date:
        result["suspension_start_date"] = suspension_date.group(1).strip()
    
    #three files still having issues with this. Unsure why. Manually updated in output
    if suspension_duration:
        result["suspension_duration_days"] = int(suspension_duration.group(1).strip())
        
    return result




def process_directory(dir_path):
    ''' function to iterate directory, read files, and set up dataframe
            Args:
                dir_path (str): file path to directory
    '''
    
    data = []
    
    ignore_value = "warning"
    
    #iterate over items in folders in a root directory
    for root, dirs, files in os.walk(dir_path):
        for filename in files:
            file_path = os.path.join(root, filename)
            ext = os.path.splitext(filename)[1].lower()
            
            if ignore_value not in filename:
                #try:
                if ext == ".txt":
                    with open(file_path, "r", encoding="utf-8") as f:
                        text = f.read()
                elif ext == ".pdf":
                    text = txt_from_pdf(file_path)                    
                elif ext == ".docx":
                    text = docx2txt.process(file_path)
                else:
                    continue
                    
                parsed = parse_letter(text)
                
                #create output using processed info
                row = {
                    "email_name" : os.path.basename(root),
                    "document_name" : filename,
                    "recipient_name" : parsed["recipient_name"],
                    "address" : parsed["address"],
                    "dates" : "; ".join(parsed["dates"]),
                    "general_laws" : "; ".join(parsed["general_laws"]),
                    "license_actions" : "; ".join(parsed["license_actions"]),
                    "suspension_start_date" : parsed["suspension_start_date"],
                    "suspension_duration_days" : parsed["suspension_duration_days"]
                    }
                
                #slap it all together
                data.append(row)
                    
                # except Exception as e:
                #     print(f"Error processing {filename}: {e}")
    
    #put it in a dataframe (table)
    df = pd.DataFrame(data)
    return df


#------------------------------
#Testing

dir_working = r"C:\Users\TAM4027\Documents\_MSLC\attach2"

#convert doc to docx for processing
doc_to_docx(dir_working)

df = process_directory(dir_working)

#save output as a csv to working directory (where python file is saved)
df.to_csv("parsed_email_sumamary.csv", index=False)
print("\nSaved parsed_email_sumamary.csv\n")