"""
    @author: Himani Bhatia, himanibhatia0907@gmail.com
    Date:    01/05/2015

    Description:
                This file extracts information from given sample PDF files 
                and puts the desired reult into two seperate csv files.
"""
from cStringIO import StringIO
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfdevice import PDFDevice
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.layout import LAParams, LTTextBox, LTTextLine,LTFigure,LTRect,LTImage,LTAnno
from pdfminer.converter import PDFPageAggregator
import copy
import os
import re 
import csv

def dataExtract(path):
    # create the header for file 1
    list1 = ['PDF_Name','Commentary','Commentary(text)']

    # pass it to csv file
    myfile  = open('File 1.csv',"wb") 
    wr      = csv.writer(myfile, quoting=csv.QUOTE_ALL)
    wr.writerow(list1)
    myfile.close()

    # create the header for file 2
    list2 = ['PDF_Name','Volatility (0/1)','Volatility (text)','Volatility (score)','Alpha (0/1)','Alpha (text)','Alpha (score)']

    # pass it to csv file
    with open('File 2.csv',"wb") as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
        wr.writerow(list2)

    # List of PDF files to be extracted
    filenames = [file for file in os.listdir(path) if re.match(r'^.+\.pdf$', file,re.I)]
    for file in filenames:
        print file
        global tb 
        
        # list for storing textboxes
        tb = [] 

        # convert the pdf to text file here
        convert(file,path)

        # Initialize local variables
        comment_text    = ""
        vol_text        = ""
        alpha_text      = ""
        alpha_score     = ""
        vol_score       = ""
        comment_flag    = 0
        bool_flag       = True
        bool_flag_2     = True
        vol_flag        = 0
        alpha_flag      = 0
        vol_alpha_score = []
        al_flag         = True
        iterator1       = 0
        iterator2       = 0

        # iterate through every textbox
        for i in range(len(tb)):
            # iterate through every textline in the textbox
            for j in range(len(tb[i])):
                # check comments in file
                matchObj_commentary = re.search( r'Fund Manager\'s Commentary|Fund Manager\'s Comment|Fund Manager\'s Report|Manager\'s Comment|Fund Commentary|Commentary|Management Comments|Comments|Commentary', tb[i][j], re.M|re.I)
                if matchObj_commentary and bool_flag:
                    comment_flag    = 1
                    bool_flag       = False

                    # if the given comment is 'Fund Manager's Comment'
                    if re.search( r'Fund Manager\'s Comment', tb[i][j], re.M|re.I):
                        comment_text = 'Fund Manager\'s Comment \n'
                        if (i+106) in tb:
                            for x in range(len(tb[i+106])):
                                comment_text = comment_text + tb[i+106][x]
                    
                    # if the given comment is 'Manager's Comment'
                    elif re.search(r'Manager\'s Comment', tb[i][j], re.M|re.I):
                        comment_text = tb[i+1][j]
                    
                    # any other comments
                    else:
                        comment_text = tb[i][j]

                # block to check volatility keyword in file
                if re.search('Volatility', tb[i][j], re.M|re.I):
                    vol_flag = 1

                    # if volatlity is a single word in line, then check for a numerical value in adjacent text box OR check the next line in same textbox
                    if re.match('Volatility *\\n+',tb[i][j],re.I) :
                        if bool_flag_2:
                            if(re.match('-?\d+\.?\d+%? *',tb[i+1][j].encode('utf-8'))):
                                bool_flag_2     = False
                                vol_text        = tb[i][j] + tb[i+1][j]
                                vol_alpha_score = re.findall('-?\d+\.?\d+%? *',tb[i+1][j].encode('utf-8'))
                                vol_score       = vol_alpha_score[0]

                                # Check for the alpha keyword if volatility keyword is found
                                if re.search('Alpha', vol_text, re.M|re.I):
                                    alpha_flag  = 1
                                    alpha_text  = vol_text
                                    alpha_score = vol_alpha_score[1]
                                    al_flag     =  False
                            
                            # if 'volatility' keyword is in the next line of same textbox
                            elif re.search('-?\d+\.?\d+%? *',tb[i][j]):
                                bool_flag_2 = False
                            
                            else:
                                vol_text    = tb[i][j]

                    # if doesnt find the best match (score horizon), then it should pick up all matched text areas for the keyword 'volatility'
                    else:
                        iterator1   = iterator1 + 1
                        vol_text    = vol_text  + '\n======MATCH:'+str(iterator1) + '======\n' + tb[i][j]
                
                # block to check alpha keyword in file
                if(al_flag and re.search( r'Alpha', tb[i][j], re.M|re.I)):
                    alpha_flag = 1

                    # 
                    if re.match('Alpha *\\n',tb[i][j],re.I) :
                        alpha_score = re.findall('-?\d+\.?\d+%? *',tb[i+1][j].encode('utf-8'))
                        alpha_text  = tb[i][j]
                        al_flag     = False
                    
                    # if doesnt find the best match (score horizon), then it should pick up all matched text areas for the keyword 'alpha'
                    else:
                        iterator2   = iterator2 + 1
                        alpha_text  = alpha_text + '\n======MATCH:' + str(iterator2) + '======\n' + tb[i][j]

        # insert row in File 1 
        list1 = [file, comment_flag, comment_text.encode("utf-8")]
        with open('File 1.csv', "a") as myfile:
            wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
            wr.writerow(list1)

        # insert row in File 2
        vol_text    = vol_text.encode("utf8")
        alpha_text  = alpha_text.encode("utf8")
        list2       = [file, vol_flag, vol_text,vol_score,alpha_flag,alpha_text,alpha_score]
        with open('File 2.csv', "a") as myfile:
            wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
            wr.writerow(list2)

def lt(layout):
    tl = [""] # list for storing text lines

    # iterate through textboxes and textlines
    for lt_obj in layout:
        if isinstance(lt_obj, LTTextBox):
            tb.append(copy.copy(lt(lt_obj)))
        if isinstance(lt_obj, LTTextLine):
            txt = lt_obj.get_text()
            if txt.isspace():
                tl.append("")
            else:
                tl[-1] = tl[-1] + txt
    return tl

def convert(my_file,path):

    base_path = path

    my_file = os.path.join(base_path + "/" + my_file)
    password        = ""
    extracted_text  = ""

    # Open and read the pdf file in binary mode
    fp  = open(my_file, "rb")

    # Create parser object to parse the pdf content
    parser = PDFParser(fp)

    # Store the parsed content in PDFDocument object
    document = PDFDocument(parser, password)
        
    # Create PDFResourceManager object that stores shared resources such as fonts or images
    rsrcmgr = PDFResourceManager()

    # set parameters for analysis
    laparams = LAParams()

    # Create a PDFDevice object which translates interpreted information into desired format
    # Device needs to be connected to resource manager to store shared resources
    # Extract the decive to page aggregator to get LT object elements
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)

    # Create interpreter object to process page content from PDFDocument
    # Interpreter needs to be connected to resource manager for shared resources and device 
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    # Ok now that we have everything to process a pdf document, lets process it page by page
    for page in PDFPage.create_pages(document):

        # As the interpreter processes the page stored in PDFDocument object
        interpreter.process_page(page)

        # The device renders the layout from interpreter
        layout = device.get_result()

        # Out of the many LT objects within layout, we are interested in LTTextBox and LTTextLine
        lt(layout)

    #close the pdf file
    fp.close()

# Uncomment the below line and give path of folder containing PDF files to run the script
dataExtract("D:\Analytics\Internships\Task\Task\PDF")
