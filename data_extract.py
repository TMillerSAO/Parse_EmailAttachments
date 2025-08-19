# -*- coding: utf-8 -*-
"""
Created on Wed Jul 30 14:11:13 2025

@author: TAM4027
"""

import zipfile
import os
import extract_msg
import re
import sys

#define working directory

dir_working = r"C:\Users\TAM4027\Documents\_MSLC"
# dir_working = "\\\\Sao-bos-fp01\\public\\groups\\AO_DAU\\DAW\\Travis Miller\\Active Audits\\\
# 2025-0089-3S Massachusetts Lottery Commission (MSLC)\\Audit Documents\\Data\\Emails"

def log_errors(log_file, msg, attachment_filename, error):
    
    '''
    Logs any errors to a file for review
    
    Args:
        log_file: (str) text error file name
        msg: (str) message file containing error
        attachment_filename: (str) attachment name in msg that caused error
        error: (str) error string
    '''
    
    with open(log_file, "a", encoding = "utf-8") as log:
        log.write(f"[{msg}] Attachment: {attachment_filename} - Error: {error}\n")


#clean up null characters        
def sanatize_filename(filename):
    '''
    Cleans up inputted string of null characters, null bits, or illegal characters.
    Returns cleaned string or attachment.bin if NONE
    
    Args:
        filename: (str) string to be cleaned
    '''
    #filename must be a string
    if not isinstance(filename, str):
        filename = str(filename)
    
    #replaces null bits and illegal characters in filename
    filename = filename.replace("\x00", '')
    filename = re.sub(r'[<>:"/\\|?*\x00-\x1F"]', '', filename).strip()
    
    #if no filename, return attachment.bin (binary file)
    return filename.strip() or "attachment.bin"



def extract_attachments(msg_path,
                        dir_output,
                        level = 0, 
                        log_file = "extraction_error.log"):
    '''
    Extracts attachments from .msg or .zip.

    Args:

        msg_path : (str) full filepath of files from original .zip. Should be .msg, otherwise investigate file type
        dir_output : (str) full filepath of where attachments should be saved to 
        level : (int) ignore, part of a test where if a .msg contained a .msg, it would perform recursion. DOES NOT CURRENTLY WORK
            The default is 0.
        log_file : (str) name of file for errors review
            The default is "extraction_error.log".

    Returns None. This is only to extract attachments and save them elsewhere

    '''
    
    #parses .msg items
    msg = extract_msg.Message(msg_path)
    msg_filename = os.path.basename(msg_path).strip()
    
    print(f"\nProcessing email: {msg_filename}")
       
    if msg_filename.endswith('.msg'):
       
        # Create a subfolder for attachments of each email
        # if one already exists for an email, skip
        dir_cur_output = os.path.join(dir_output, os.path.splitext(msg_filename)[0].strip())
        if not os.path.exists(dir_cur_output):
            os.makedirs(dir_cur_output, exist_ok=True)
        else:
            print(f"{dir_cur_output} already processed. ***** Checking attachments....")     

        # Extract and save each attachment
        for attachment in msg.attachments:
            try:
                #obtain and clean attachment name ex: xyz.doc
                #use the attachment's name otherwise name it "unknown"
                attachment_filename = sanatize_filename((
                    attachment.longFilename if attachment.longFilename 
                    else (attachment.shortFilename if attachment.shortFilename 
                    else "unknown")
                    ))
                
                attachment_path = os.path.join(dir_cur_output, attachment_filename)
                
                
                #skips image attachments
                if attachment_filename.lower().startswith("image"):
                    print(f"   Skipping image attachment: {attachment.longFilename}")
                    continue                    
    
                #handle potential dupes
                if os.path.exists(attachment_path.strip()):
                    print(f"   {attachment_filename} already processed. {attachment_path}")
                    continue
    
                else:
                    #saves attachment to email sub-folder of attachment directorty
                    print(f"   Saving new item {attachment_filename} to {dir_cur_output}")
                    
                    #if attachment is a message object, append .msg and save
                    if attachment.data.__class__.__name__ == "Message":
                        if not attachment_filename.lower().endswith(".msg"):
                            attachment_filename += ".msg"
                            #print(attachment_filename)
                            attachment_path = os.path.join(dir_cur_output, attachment_filename)
                            
                        ###BROKEN### --> cannot get to nested emails
                        #ended up doing manually (only 3 files with =< 2 nested)
                        attachment.save(customPath = attachment_path)
                            
                            #recursive .msg attachments
                        extract_attachments(attachment_path, dir_cur_output, level+1)
                    
                    #if attachment is a .zip, redo unzip_folder and save insides
                    #Each .zip attachment in this case contained .doc or .pdf so don't have to redo processing
                    elif attachment_filename.endswith(".zip"):
                        try:
                            #needs to be bytes
                            #save .zip
                            with open(attachment_path, "wb") as f:
                                f.write(attachment.data)
                        except Exception as e:
                            print(f" Error processing {attachment_filename}: {e}...Trying second save option...")
                            attachment.save(customPath = attachment_path)
                          
                        #current output directory will be save location for attachment extract
                        #"email name" in result will be email containing the zip file. 
                            #i.e. nested attachments will be attributed to highest email    
                        unzip_folder(attachment_path,
                                     dir_unzipped = dir_cur_output,
                                     att_process_needed = False)
                    
                    else:
                        #see if there are null bytes in path
                        if '\x00' in str(attachment_path):
                            #if so, kill program cause that needs to be addressed
                            sys.exit()
                        else:
                            try:
                                #needs to be bytes
                                #save attachment
                                with open(attachment_path, "wb") as f:
                                    f.write(attachment.data)
                            except Exception as e:
                                print(f" Error processing {attachment_filename}: {e}...Trying second save option...")
                                attachment.save(customPath = attachment_path)
                        
                    print(f"   Successfully extracted: {attachment.longFilename} from {msg_filename}")
               
            except Exception as e:
                print(f"Error processing {attachment_filename}: {e} \n")
                log_errors(log_file, msg_filename, attachment_filename, e)
   
        print(f"Attachment extraction for {msg_filename} complete.")
    else:
        #check msg_filename extention
        print(f"File {msg_filename} does not have extention .msg")



# Process each extracted email file 
def process_folder(dir_unzipped,
                   dir_attachment = dir_working + "\\attach"):
    
    '''
    Extracts and reads attachments from emails.
    
    Args:
        dir_unzipped (str): The directory of extracted emails. 
            Defaults to/creates "unzipped" in the dir_working directory following unzip_folder() pass off
            
        dir_attachment (str): directory to dump attachments in emails
            Defaults to/creates "attachments" in the dir_working directory.
    '''
    
    # Create the output folder if it doesn't exist
    if not os.path.exists(dir_attachment):
        os.makedirs(dir_attachment, exist_ok=True)
    
    # Iterate through each file in the email folder
    for msg_filename in os.listdir(dir_unzipped):
        msg_path = os.path.join(dir_unzipped, msg_filename)
    
        extract_attachments(msg_path, dir_attachment)



        
def unzip_folder(ZipPath,
                 dir_unzipped = dir_working + "\\unzipped",
                 att_process_needed = True,
                 dir_attachment = dir_working + "\\attach"
                 ):
    """
    Extracts files from a zip archive

    Args:
        ZipPath (str): The path to the zip file containing email files.
        
        dir_unzipped (str): The directory of extracted emails. 
            Defaults to/creates "unzipped" in the dir_working directory.
            
        att_process_needed (T/F): If True, dir_unzipped is passed to  
            process_folder()
            
        dir_attachment (str): directory to dump attachments in emails. 
            Only used to pass to process_folder()
            Defaults to/creates "attachments" in the dir_working directory.
    """
    #clear error log
    open("extraction_error.log", "w").close()
    
    #creates directory to store unzipped files
    if not os.path.exists(dir_unzipped):
        os.makedirs(dir_unzipped)

    with zipfile.ZipFile(ZipPath, 'r') as zip_ref:
        #Reads each member in zip file
        for member in zip_ref.infolist():
            target = os.path.join(dir_unzipped, member.filename)
        
            #Checks if member is already in target folder
            if os.path.exists(target):
                #if yes, skip
                print(f"Already processed {target}.")
            else:
                #if not, extract member
                #dir_unzipped will be save location for attachment extract
                zip_ref.extract(member, dir_unzipped)
                
            
    if att_process_needed:
        #flag to process a folder. Zips containing only .doc/.pdf do not need to be processed
        process_folder(dir_unzipped, dir_attachment)
            


#-------------------------------------------------------------
#Testing

zip_file_1 = dir_working + "\\18771379.zip"
zip_file_2 = dir_working + r'\Tobacco Enforcement Letters_87.zip'

#test extraction function
unzip_folder(ZipPath = zip_file_1,
               dir_unzipped = dir_working + "\\unzipped2",
               dir_attachment = dir_working + "\\attach2")







