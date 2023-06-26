#run on dev (3.9, 64-bit) environment (tho VS may not support)
import os
import datetime
import csv
import pandas as pd



MoFolder = 'C:\\Users\\cara.fagerholm\\Documents\\Incubator Mgmt\\MonthlyInvoicing'

dfAte=pd.read_excel(MoFolder+r"\Ate EQ reservation Report - 7_30-6.xlsx", sheet_name='Ate EQ reservation Report - 7_3', skiprows=1)
dfAte.columns


ListDirCur = os.listdir(MoFolder)

#read in Ateios smartsheet outline into dataframe(? if can easily edit, I assume)
    #build reference data table based on ease of #1, for #2
#read in reference data table/db of rates and restrictions, 
#set variables of: start and end peak time, total hrs/mo
#add a column for peak usage, add a col for non-peak usage
#compare total hrs/mo with summed peak/non-peak
#subtract what is overage
#multiply rates against overage eq/time usage
#sum
#print formatting txt file
    #headers
    #each eq used in sheet, overage amt, rate, cost




with open("C:\\Users\\cara.fagerholm\\Documents\\Incubator Mgmt\\Summary.csv", 'w') as f:
    for Indicie, Arbin in enumerate(ListDirCur):

        # current folder path
        current_folder = os.path.join(curfolder, Arbin)
        print(current_folder)

        list_dir = os.listdir(current_folder)

        for index, dataFolder in enumerate(list_dir):
            print(dataFolder[0:8])
            if dataFolder[0:8] != "Complete":
                #writer = csv.writer(file)
                #writeLine = os.path.join(current_folder, dataFolder)
                #writer.writerow(writeLine)
                f.write(os.path.join(current_folder, dataFolder))
                f.write('\n')
        
#test tracking file
with open("C:\\Users\\cfagerholm\\TestTracking.tsv", 'wt') as g:
    for Indicie, Arbin in enumerate(ListDirCur):

        current_folder = os.path.join(curfolder, Arbin)
        list_dir = os.listdir(current_folder)
        for index, dataFolder in enumerate(list_dir):
            if dataFolder[0:8] != "Complete":
                g.write(dataFolder[0:4])
                g.write('\t')
                g.write('\t')
                g.write(dataFolder)
                g.write('\t')
                g.write('\t')
                g.write('\t')
                g.write('\t')
                g.write('\t')
                g.write(Arbin)
                g.write('\t')
                chNumsList = []
                list_ch = os.listdir(os.path.join(current_folder, dataFolder))
                for i, chFilename in enumerate(list_ch):
                    ch = chFilename[-18:-16]
                    if chFilename[-18:-17] == "_":
                        ch = chFilename[-17:-16]
                    if ch.isdigit() == False:
                        ch = chFilename[-12:-10]
                        if chFilename[-12:-11] == "_":
                            ch = chFilename[-11:-10]
                    chNumsList.append(ch)
                #g.write(f'{chNumsList}')
                #print(f'{chNumsList}')
                print(', '.join(chNumsList))
                g.write(', '.join(chNumsList))
                g.write('\t')
                now = datetime.now()
                today = now.strftime("%m/%d/%Y")
                g.write(today)
                g.write('\n')
        g.write('\n')



