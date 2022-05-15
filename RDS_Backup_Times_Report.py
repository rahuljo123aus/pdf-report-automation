# AUTHOR: RAHUL JOSEPH
# SCRIPT: PYTHON SCRIPT TO GENERATE REPORT


# This report is generated to show the rds backup times for in-house auditing 


# snippet of the backup data collected from gdrive
#
# date, start time, end time , duration (end-start),env (production or test)




from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import os, sys ,math
from os import listdir
import datetime
import pandas as pd
import numpy as np
import dataframe_image as dfi
import matplotlib.pyplot as plt
from fpdf import FPDF
from pandas.plotting import table
import plotly.figure_factory as ff
import plotly.graph_objects as go

# Ignoring pandas loc function warnings
pd.set_option('mode.chained_assignment',None)


# Creating Access Token to connect to Google Drive API
# If Access Token is Empty, load new tokens
# Else refresh the tokens
gauth = GoogleAuth()
# Try to load saved client credentials
gauth.LoadCredentialsFile("mycreds.txt")
if gauth.credentials is None:
    # Authenticate if they're not there
    gauth.LocalWebserverAuth()
elif gauth.access_token_expired:
    # Refresh them if expired
    gauth.Refresh()
else:
    # Initialize the saved creds
    gauth.Authorize()
    print("\nGoogle Authentication Successful!")

# Save the current credentials to a file
gauth.SaveCredentialsFile("mycreds.txt")

# Initialize Google Drive API
drive = GoogleDrive(gauth)

# Report Data collected from Google Drive Folder ID
parent_folder_id = ' original folder '



# Function to download files from Google Drive
def downloadfiles(fileName):
    file_list = drive.ListFile({'q': "'%s' in parents and trashed=false" % parent_folder_id}).GetList()

    for file in file_list:
        if fileName == file['title']:
            file.GetContentFile(fileName)
            print(f"Downloading {fileName}...")
            

# Function to list out all the files in the Google Drive Folder
def listfiles():
    file_list = drive.ListFile({'q': "'%s' in parents and trashed=false" % parent_folder_id}).GetList()

    file_names = []

    for file in file_list:
        
        if file['lastModifyingUserName'] == 'me': 
            file_names.append(file['title'])

    return file_names


def extractlatestdata(fileName,absolute_path,month,year):
    customer_name = fileName.split('.')[0].lower()

    df = pd.read_csv(fileName,names=["Date", "Start", "End", "Value" , "Env"])
  
      

    for index,row in dfprd.iterrows():
        date = datetime.datetime.strptime(str(row['Date']).replace(",",""), '%A %d %B %Y').strftime('%m/%d/%Y')
        dfprd.loc[index, 'Date'] = date
        value = round(row['Value'],2)
        dfprd.loc[index, 'Value'] = value

    
    dfprd['Year'] = pd.DatetimeIndex(dfprd['Date']).year
    dfprd['Month'] = pd.DatetimeIndex(dfprd['Date']).month 
    

    
    for index,row in dfprd.iterrows():
        date = datetime.datetime.strptime(str(row['Date']), '%m/%d/%Y').strftime('%d %b %Y')
        dfprd.loc[index, 'Date'] = date
        
        
    
    dfprd = dfprd[['Date','Start','End','Value','Env']]
    dfprd1 = dfprd.copy()

 
    colorscale = [[0, '#29ade6'],[.5, '#ffffff'],[1, '#f5f5f5']]
    fig =  ff.create_table(dfnpe1,colorscale=colorscale)
    fig.update_layout(font=dict(family= "Arial"))
    fig.update_layout(autosize=True,font_size = 16)
    fig.write_image(absolute_path + r'\department\images\{0}_dfnpe.png'.format(customer_name.upper()),scale=2)

    

    # calling the function to plot and save the plots as images.
    #     
    generate_matplotlib_stackbars(dfprd,absolute_path + r"\department\images\{0}_dfprd_barchart.png".format(customer_name.upper()),"#66e312")


def generate_matplotlib_stackbars(df, filename,colour):
    

    length = len(df['Date'].values)
    x_pos = [((x * 2) + 5.5) for x in range(length)]

        
    # Create subplot and bar
    fig, ax = plt.subplots()

    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    bars = ax.bar(x_pos, df['Value'].values, color=colour,align="center",width=1.0) 

    # Set Title
    # ax.set_title('{0} {1} Database Backup Time ({2})\n\n'.format(customer,mode,rds), fontweight="bold")

    # Set xticklabels
    plt.xticks(x_pos,df['Date'].values,rotation = 90, fontsize=7)
    # ax.set_xticklabels(df['Date'].values, rotation=90,fontsize = 8)

    # set yticklabels
    plt.yticks(np.arange(0, 6, 1),fontsize=7)
    plt.ylim(top = 4.85)

    # access the bar attributes to place the text in the appropriate location
    for bar in bars:
        yval = bar.get_height()
        if yval > 4.85:
            plt.text(bar.get_x() - 0.40, 4.85 + 0.07, yval,fontsize=7,rotation=90)
        else:
            plt.text(bar.get_x() - 0.40, yval + 0.07, yval,fontsize=7,rotation=90)
 
    
    # Set ylabel
    ax.set_ylabel('Backup Time (in mins)', fontsize=7) 

    # Set xlabel
    # ax.set_xlabel('\nDate', fontsize=7) 

    # Save the plot as a PNG
    plt.savefig(filename, dpi=600, bbox_inches='tight', pad_inches=0)
    
    # Display the figure
    # plt.show()

    # Closing the matplotlib object
    plt.close()

def create_letterhead(pdf, WIDTH,absolute_path):
    pdf.image(absolute_path+r"\templates\page_header2.png", 0, 0, WIDTH)


def generatepdf(customer,absolute_path,month,year):
    
    WIDTH = 210

    # Create PDF
    pdf = FPDF('P', 'mm', (WIDTH, 320)) 

    print("Creating First Page")
    '''
    First Page of PDF
    '''
    # Add Page
    pdf.add_page()

                # Page 1 Header

    # Add lettterhead and title
    pdf.ln(4)
    create_letterhead(pdf, WIDTH,absolute_path)

                # Page 1 Month Year

    # Set font
    pdf.set_font('Arial', 'B', 10)
    # Move to 8 cm to the right
    pdf.cell(80)
    # Centered text 
    pdf.cell(25, 15, str(month)+' '+str(year) , 0, 0, 'C')
    # linebreak
    pdf.ln(8)

                # Page 1 Heading 1

    # Set font
    pdf.set_font('Arial', 'B', 10)
    # Move to 8 cm to the right
    pdf.cell(80)
    # Centered text 
    pdf.cell(20, 30, 'Production', 0, 0, 'C')
    # line break
    pdf.ln(20)

                # Page 1 Table

    # Add table
    pdf.cell(20)
    pdf.image(absolute_path + r"\department\images\{0}_dfprd.png".format(customer), w=150,h=150)
    # Line Break
    pdf.ln(7)

                # Page 1 Heading 2

    # Set font
    pdf.set_font('Arial', 'B', 10)
    # Move to 8 cm to the right
    pdf.cell(80)
    # Centered text 
    pdf.cell(30, 15, 'Production Backup Time', 0, 0, 'C')
    # line break
    pdf.ln(15)

                # Page 1 Visualisation    

    # Move right
    # pdf.cell(5)
    # Add the generated visualisation to the PDF
    pdf.image(absolute_path + r"\department\images\{0}_dfprd_barchart.png".format(customer), w=180, h=70, type='png')
    # line break
    pdf.ln(10)
                # Page 1 Numbering the page

    # Add Page number
    # Go to next line
    pdf.set_y(-33)
    # Select Arial italic 8
    pdf.set_font('Arial', 'I', 8)
    # Print centered page number
    pdf.cell(0, 10, 'Page 1', 0, 0, 'C')
    

    print("Finished First Page")

    ##############################
 

    # Generate the PDF
    pdf.output(absolute_path + r"\backuppdfReports\{0} Daily Backup Report {1} {2}.pdf".format(customer,month,year), 'F')

if __name__=='__main__':

    currentdate = datetime.datetime.now().strftime('%Y-%m-%d')
    firstdate = datetime.datetime.now().replace(day=1).strftime('%Y-%m-%d')

    # to run manually, uncomment the below line
    currentdate =  firstdate    

    if currentdate != firstdate:
        print("Not the first day of the month...\nSkipping ....")
        sys.exit(-1)
    
    print("\nBACKUP REPORT GENERATOR STARTED!\n")

    scriptstarttime = datetime.datetime.now()
    scriptrunday = scriptstarttime.strftime("%A").lower()[:3]
    last_month_date = scriptstarttime.replace(day=1) - datetime.timedelta(days=1)
    last_month_full_name = last_month_date.strftime("%B")
    last_month_num = last_month_date.strftime("%m").lstrip("0")
    last_month_year = last_month_date.strftime("%Y")


    month = int(last_month_num)
    year = int(last_month_year)


    absolute_path = os.getcwd()
   
    filenames = listfiles()
    print(filenames)
    
    if str(os.getcwd()) != str(absolute_path): 
        print("Switching working directory!")
        print("FROM: {0}".format(os.getcwd()))
        print("TO: {0}".format(absolute_path))
        os.chdir(absolute_path)
    else:
        print("Working Directory: {0}".format(os.getcwd()))

    if os.path.exists(absolute_path + r"\backuppdfReports") is False:
        os.mkdir("backuppdfReports")

    if os.path.exists(absolute_path + r"\department") is False:
        os.mkdir("department")

    if os.path.exists(absolute_path + r"\department\images") is False:
        os.mkdir("images")
    
    os.chdir(absolute_path + r"\department")

    for eachFile in filenames:
        downloadfiles(eachFile)
    

    # Empty the backupReport folder
    # below lines of code would empty the folder
    print("\nEmptying backReports folder...\n")
    backupreport_path = absolute_path + r"\backuppdfReports"
    for filename in os.listdir(backupreport_path):
        file_path = os.path.join(backupreport_path, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))


    for eachFile in listdir(absolute_path + r"\department"):
        if eachFile != 'images':
            customer = eachFile.split('.')[0].upper()
            print(f"\nBuilding Images for {customer} Report!")
            extractlatestdata(eachFile,absolute_path,month,year)
            print(f"Generating PDF Report for {customer}")
            generatepdf(customer,absolute_path,last_month_full_name,last_month_year)
    
      