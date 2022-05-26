from datetime import datetime
import fitz
from os import chdir, getcwd
from glob import glob as glob
import re
import pandas as pd

def main():
    hcNumRe = r'(?<=Healthcare plan #: )[0-9]+'
    famNumRe = r'(?<=Family member #: )[0-9]+'

    df = pd.read_excel(r'Patient list.xlsx', sheet_name='Visits')
    df = df.set_index('Health care plan #')

    get_curr = getcwd()
    directory = 'FILES'
    newName = ''

    chdir(directory)
    
    pdf_list = glob('*.pdf')
    chdir(get_curr +'/NEW_FILES')

    for pdf in pdf_list:
        with fitz.open(get_curr + '/FILES/' + pdf) as pdf_object:
            for count, page in enumerate(pdf_object):

                #Grabbing the pdf text for searching
                pdfText = page.get_text()

                #Getting the attributes to use to search in the excel file
                famN = grabPageInfo(pdfText, famNumRe)
                hcN = grabPageInfo(pdfText, hcNumRe)

                #Query the excel file for related row (not a great way to search)
                tempDF = df.loc[df['Family #'] == int(famN)].loc[int(hcN)]

                #Creating new PDF file name from the excel file using the following columns:
                #"Appointment date", "Patient #", "Appointment code"
                newName = str(datetime.strptime(str(tempDF['Appointment date']), '%Y-%m-%d %H:%M:%S').date()) + '_' + str(tempDF['Patient #']) + '_' + str(tempDF['Appointment code']) + '.pdf'

                #Putting the documents 
                newDoc = fitz.open()
                newDoc.insert_pdf(pdf_object, from_page=count, to_page=count)
                newDoc.save(newName)
                
def grabPageInfo(pdf_text, regexC):
    return re.search(regexC, pdf_text).group().strip()

if __name__ == "__main__":
    main()