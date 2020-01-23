import os
import pandas as pd
from statistics import mean
def prompt():
    # folder_path = input('Please enter the folder: ').strip()
    # while os.path.isdir(folder_path)== False:
    #     print('This folder does not exist, please try again.')
    #     folder_path = input('Please enter the folder: ').strip()

    folder_path = 'C:\\Users\\ANDREW\\Desktop\\IC50'
    return folder_path

def open_file_IC50(folder_path):
    for file in os.listdir(folder_path):
        print('Reading file:', file)
        if "Results" in file:
            pass
        else:
            file_save = file.split('_')
            file_save = file_save[0] + '_' +file_save[1] +'_' + 'Results.xlsx'
            file_save = os.path.sep.join([folder_path, file_save])

            writer = pd.ExcelWriter(file_save, engine='xlsxwriter') #Creates the file to write to

            file_path = os.path.sep.join([folder_path,file])
            xls = pd.ExcelFile(file_path)

            selectivity_df = pd.DataFrame()
            for sheet in xls.sheet_names:
                if sheet == '20-HETE-d6' or sheet == 'Component':
                    pass
                else:
                    df = pd.read_excel(file_path, sheet_name = sheet)
                    new_df = organize_sheet(df)
                    new_df = analyze_sheet(new_df, sheet, file)
                    new_df.to_excel(writer, sheet_name=sheet)
                if sheet == '20-HETE-d6' or sheet == 'Component':#Purpose of this is to generate reg curves of the other metabolites and IC50
                    pass
                else:
                    if 'Chen' in file:#Chenxiao specific, 12 points on IC50 curve
                        #write_csv_chenxiao(new_df,folder_path, file, sheet)
                        #If needed can create IC50 curves here, but need to create the function for it. Need also the points on the curve
                        pass
                    elif len(new_df['Filename']) > 67:#2 IC50s at once
                        write_csv_option2(new_df, folder_path, file, sheet)
                    else:#One IC50, 10 points on curve
                        write_csv_option1(new_df, folder_path, file, sheet)
                if sheet == '20-HETE-d6' or sheet == 'Component' or sheet == '20-HETE' or sheet == '15-HETE' or sheet == '12-HETE':
                    pass
                else:
                    if 'Chen' in file:#Chenxiao specific, 12 points on IC50 curve
                        selectivity_df[sheet] = new_df['10-fold'][:42]
                    elif len(new_df['Filename']) > 67:#2 IC50s at once
                        selectivity_df[sheet] = new_df['10-fold'][:66]
                    else:#One IC50, 10 points on curve
                        selectivity_df[sheet] = new_df['10-fold'][:36]
            
            if 'Chen' in file:#Chenxiao specific, 12 points on IC50 curve
                selectivity_df = organize_selectivity_chenxiao(selectivity_df)
            elif len(new_df['Filename']) > 67:#2 IC50s at once
                selectivity_df = organize_selectivity_option2(selectivity_df)   
            else:#One IC50, 10 points on curve
                selectivity_df = organize_selectivity_option1(selectivity_df)           

            selectivity_file_name = file.split('_')[0] + '_' + 'Selectivity.xlsx'
            selectivty_file_path = os.path.sep.join([folder_path, selectivity_file_name])
            writer_selectivity = pd.ExcelWriter(selectivty_file_path, engine = 'xlsxwriter')
            selectivity_df.to_excel(writer_selectivity, sheet_name = 'Selectivity')
            writer_selectivity.save()
            writer.save()

def organize_selectivity_option2(selectivity_df):

    list_sum = []
    for i in range(2,68):#Adds the sum for each sample
        list_sum.append('=sum(B' + str(i) + ':H' + str(i) + ')')
    selectivity_df['Sum'] = pd.Series(list_sum)
    

    list_control = []
    for i in range(2,65):#Adds %control for each sample
        list_control.append("=I" + str(i) + "/average(I65:I67)")
    selectivity_df['% Control'] = pd.Series(list_control)

    list_avg = []
    for i in range(2,65,3):#Adds the average
        list_avg.append("=average(J" +str(i)+ ":J" + str(i+2) + ")")
        list_avg.append("")
        list_avg.append("")
    selectivity_df['Average'] = pd.Series(list_avg)

    list_stdev = []
    for i in range(2,65,3):#Adds the stdev
        list_stdev.append("=stdev(J" +str(i) + ":J" + str(i+2) + ")")
        list_stdev.append("")
        list_stdev.append("")
    selectivity_df['Stdev'] = pd.Series(list_stdev)

    return selectivity_df

def organize_selectivity_chenxiao(selectivity_df):

    list_sum = []
    for i in range(2,44):#Adds the sum for each sample
        list_sum.append('=sum(B' + str(i) + ':H' + str(i) + ')')
    selectivity_df['Sum'] = pd.Series(list_sum)
    

    list_control = []
    for i in range(2,41):#Adds %control for each sample
        list_control.append("=I" + str(i) + "/average(I41:I43)")
    selectivity_df['% Control'] = pd.Series(list_control)

    list_avg = []
    for i in range(2,38,3):#Adds the average
        list_avg.append("=average(J" +str(i)+ ":J" + str(i+2) + ")")
        list_avg.append("")
        list_avg.append("")
    selectivity_df['Average'] = pd.Series(list_avg)

    list_stdev = []
    for i in range(2,38,3):#Adds the stdev
        list_stdev.append("=stdev(J" +str(i) + ":J" + str(i+2) + ")")
        list_stdev.append("")
        list_stdev.append("")
    selectivity_df['Stdev'] = pd.Series(list_stdev)

    return selectivity_df
    
def organize_selectivity_option1(selectivity_df):

    list_sum = []
    for i in range(2,38):#Adds the sum for each sample
        list_sum.append('=sum(B' + str(i) + ':H' + str(i) + ')')
    selectivity_df['Sum'] = pd.Series(list_sum)
    

    list_control = []
    for i in range(2,32):#Adds %control for each sample
        list_control.append("=I" + str(i) + "/average(I35:I37)")
    selectivity_df['% Control'] = pd.Series(list_control)

    list_avg = []
    for i in range(2,32,3):#Adds the average
        list_avg.append("=average(J" +str(i)+ ":J" + str(i+2) + ")")
        list_avg.append("")
        list_avg.append("")
    selectivity_df['Average'] = pd.Series(list_avg)

    list_stdev = []
    for i in range(2,32,3):#Adds the stdev
        list_stdev.append("=stdev(J" +str(i) + ":J" + str(i+2) + ")")
        list_stdev.append("")
        list_stdev.append("")
    selectivity_df['Stdev'] = pd.Series(list_stdev)

    return selectivity_df


def organize_sheet(df):
    column_list = ['Component Name', 'Origin Index',
                   'Equation', 'Unnamed: 8', 'Unnamed: 14']#The columns that are wanted intially
    for column in column_list:#Create a new dataframe with only the wanted columns
        try:
            new_df = pd.concat([new_df, df[column]], axis = 1)
        except NameError:
            new_df = df[column]

    new_df.columns = new_df.iloc[2] #Adjust the column headers
    new_df = new_df[3:]#Get rid of the first 3 lines

    column1_info = []

    for cell in new_df['Filename']:#Adjust the numbers of samples so that they are 0 padded
        cell = str(cell).zfill(2)
        column1_info.append(cell)
    new_df['Filename'] = column1_info

    new_df = new_df.reset_index()#Resets the index
    new_df = new_df.drop(['index'], axis =1)#Removes the messed up index column

    for index, sample in enumerate(new_df['Filename']):
        #print(sample, len(str(sample)))
        if 'A' in str(sample) and len(str(sample))>2:
            new_sample = 'std_' + str(sample)
            new_df.at[index ,'Filename'] = new_sample
    new_df = new_df.sort_values(by=['Filename'])#Sort by the first column
    new_df = new_df[new_df.Filename != 'shutdown']
    new_df = new_df[new_df.Filename != 'startup']
    new_df = new_df[new_df.Filename != 'nan']
    new_df = new_df[new_df.Filename != 'User Name']
    new_df = new_df[new_df.Filename != 'pjo7']
    new_df = new_df[new_df.Filename != 'Created By:']
    
    new_df = new_df.reset_index()#Resets the index
    new_df = new_df.drop(['index'], axis =1)#Removes the messed up index column
    new_df.replace(['NF'],[0], inplace=True)

    return new_df

def analyze_sheet(new_df,sheet, file): 
    metabolites_mw = {'12-HETE': 320.5, '15-HETE': 320.5, '20-HETE':320.5,
                    '5,6-DiHET':338.5, '8,9-DiHET':338.5, '11,12-DiHET':338.5,'14,15-DiHET':338.5, 
                    '8,9-EET':320.5, '11,12-EET':320.5, '14,15-EET':320.5}   

    new_df['Calculated Amount'] = new_df['Amount']#Creates first calculated column
    new_df['pg/vial'] = new_df['Calculated Amount']/7.5*125
    new_df['pmol/vial'] = new_df['pg/vial']/metabolites_mw[sheet]
    new_df['pmol/min/mg'] = new_df['pmol/vial']/0.3/20
    new_df['10-fold'] = new_df['pmol/min/mg']*10

    if 'Chen' in file:#Chenxiao specific, 12 points on IC50 curve
        new_df = option_chenxiao(new_df)
    elif len(new_df['Filename']) > 67:#2 IC50s at once
        new_df = option2(new_df)
    else:#One IC50, 10 points on curve
        new_df = option1(new_df)
    return new_df

def find_prism():
    import sys
    print('Please wait, looking for Prism.')
    found = 0
    for root,dirs,files in os.walk('C:\\'):
        for file in files:
            if file == 'prism.exe':
                prism_path = os.path.sep.join([root,file])
                found = 1
                print('Prism Found.')
                break
            else:
                pass
    if found == 0:
        print('Prism was not found in C drive.')
        print('Please place Prism into your C drive.')
        input('The algorithm will now quit. Please press Enter.')
        sys.exit()
    return prism_path

def create_prism_script(folder_name): #Creates prism script in directory of python script
    cwd = os.getcwd()
    temp = 'SetPath ' + folder_name + '\n'
    path_for_prism_file = 'SetPath ' + cwd + '\n'
    script_info = [path_for_prism_file, 'Open IC50.pzfx\n', temp,
                    'ForEach *PrismIC50' +'.csv\n',
                    'Goto D\n','ClearTable\n','Import\n','Save ' + '%'+'F_IC50_Results.pzfx\n','Next\n','Beep']

    with open('prism_script.pzc','w') as file:
        for i in script_info:
            file.write(i)

def create_prism_script_regcurve(folder_name): #Creates prism script in directory of python script
    cwd = os.getcwd()
    temp = 'SetPath ' + folder_name + '\n'
    path_for_prism_file = 'SetPath ' + cwd + '\n'
    script_info = [path_for_prism_file, 'Open Reg_Curve.pzfx\n', temp,
                    'ForEach *Prism' +'.csv\n',
                    'Goto D\n','ClearTable\n','Import\n','Save ' + '%'+'F_Results.pzfx\n','Next\n','Beep']

    with open('prism_script_regcurve.pzc','w') as file:
        for i in script_info:
            file.write(i)

def run_prism_script(prism_path, script):
    from subprocess import call
    cwd = os.getcwd()
    script_path = os.path.sep.join([cwd, script])
    script_path = '@' + script_path
    call([prism_path,script_path])

def check_samples_chenxiao(new_df):
    for i in range(1,43):
        i = str(i).zfill(2)
        if str(i) in new_df['Filename'].values:
            pass
        else:
            new_df.loc[-1] = [i,0,0,0,0,0,0,0,0,0]
            new_df = new_df.sort_values(by=['Filename'])#Sort by the first column
            new_df = new_df.reset_index()#Resets the index
            new_df = new_df.drop(['index'], axis =1)#Removes the messed up index column
    return new_df

def option_chenxiao(new_df):
    #This option is for single IC50
    #The average is in column 'J'
    #Average is in 35:37

    new_df = check_samples_chenxiao(new_df)

    list_control = []
    for i in range(2,41):#Adds %control for each sample
        list_control.append("=K" + str(i) + "/average(K41:K43)")
    new_df['% Control'] = pd.Series(list_control)

    list_avg = []
    for i in range(2,38,3):#Adds the average
        list_avg.append("=average(L" +str(i)+ ":L" + str(i+2) + ")")
        list_avg.append("")
        list_avg.append("")
    new_df['Average'] = pd.Series(list_avg)

    list_stdev = []
    for i in range(2,38,3):#Adds the stdev
        list_stdev.append("=stdev(L" +str(i) + ":L" + str(i+2) + ")")
        list_stdev.append("")
        list_stdev.append("")
    new_df['Stdev'] = pd.Series(list_stdev)

    return new_df

def check_samples_option1(new_df):
    #This function checks to make sure all the samples are present in the file. Otherwise the formulas will not work out correctly
    for i in range(1,37):
        i = str(i).zfill(2)
        if str(i) in new_df['Filename'].values:
            pass
        else:
            new_df.loc[-1] = [i,0,0,0,0,0,0,0,0,0]
            new_df = new_df.sort_values(by=['Filename'])#Sort by the first column
            new_df = new_df.reset_index()#Resets the index
            new_df = new_df.drop(['index'], axis =1)#Removes the messed up index column
    return new_df

def option1(new_df):
    #This option is for single IC50
    #The average is in column 'J'
    #Average is in 35:37

    new_df = check_samples_option1(new_df)

    list_control = []
    for i in range(2,32):#Adds %control for each sample
        list_control.append("=K" + str(i) + "/average(K35:K37)")
    new_df['% Control'] = pd.Series(list_control)

    list_avg = []
    for i in range(2,32,3):#Adds the average
        list_avg.append("=average(L" +str(i)+ ":L" + str(i+2) + ")")
        list_avg.append("")
        list_avg.append("")
    new_df['Average'] = pd.Series(list_avg)

    list_stdev = []
    for i in range(2,32,3):#Adds the stdev
        list_stdev.append("=stdev(L" +str(i) + ":L" + str(i+2) + ")")
        list_stdev.append("")
        list_stdev.append("")
    new_df['Stdev'] = pd.Series(list_stdev)

    return new_df

def check_samples_option2(new_df):
    #This function checks to make sure all the samples are present in the file. Otherwise the formulas will not work out correctly
    for i in range(1,67):
        i = str(i).zfill(2)
        if str(i) in new_df['Filename'].values:
            pass
        else:
            new_df.loc[-1] = [i,0,0,0,0,0,0,0,0,0]
            new_df = new_df.sort_values(by=['Filename'])#Sort by the first column
            new_df = new_df.reset_index()#Resets the index
            new_df = new_df.drop(['index'], axis =1)#Removes the messed up index column
    return new_df

def option2(new_df):
    #This option is for single IC50
    #The average is in column 'J'
    #Average is in 65:67

    new_df = check_samples_option2(new_df)

    list_control = []
    for i in range(2,65):#Adds %control for each sample
        list_control.append("=K" + str(i) + "/average(K65:K67)")
    new_df['% Control'] = pd.Series(list_control)

    list_avg = []
    for i in range(2,65,3):#Adds the average
        list_avg.append("=average(L" +str(i)+ ":L" + str(i+2) + ")")
        list_avg.append("")
        list_avg.append("")
    new_df['Average'] = pd.Series(list_avg)

    list_stdev = []
    for i in range(2,65,3):#Adds the stdev
        list_stdev.append("=stdev(L" +str(i) + ":L" + str(i+2) + ")")
        list_stdev.append("")
        list_stdev.append("")
    new_df['Stdev'] = pd.Series(list_stdev)

    return new_df

def write_csv_option1(new_df, folder_path, file, sheet):
    concentration_list = [50, 25, 10, 2.5, 1, 0.25, 0.1, 0.025, 0.01, 0.001]

    temp_list = []
    temp = new_df['10-fold'][:30]/mean(new_df['10-fold'][33:36])*100
    for i in range(0,30,3):
        temp_list.append((concentration_list[int(i/3)], temp[i], temp[i+1], temp[i+2]))

    csv_df = pd.DataFrame(temp_list, columns = ['Concentration', '% Control', 'x', 'y'])
    if sheet == '20-HETE':
        csv_path = os.path.sep.join([folder_path, file.split('_')[0] + '_' + sheet + '_PrismIC50.csv'])
    else:
        csv_path = os.path.sep.join([folder_path, file.split('_')[0] + '_' + sheet + '_Prism.csv'])
    csv_df = csv_df.iloc[::-1] #Reverse the row order
    csv_df.to_csv(csv_path, index = False)

def write_csv_option2(new_df, folder_path, file, sheet):
    concentration_list = [50, 25, 10, 2.5, 1, 0.25, 0.1, 0.025, 0.01, 0.001]

    #This is for the first compound's csv file from the IC50 file that contains two compounds
    temp_list = []
    temp = new_df['10-fold'][:30]/mean(new_df['10-fold'][63:66])*100
    for i in range(0,30,3):
        temp_list.append((concentration_list[int(i/3)], temp[i], temp[i+1], temp[i+2]))

    csv_df = pd.DataFrame(temp_list, columns = ['Concentration', '% Control', 'x', 'y'])
    if sheet == '20-HETE':
        csv_path = os.path.sep.join([folder_path, file.split('_')[0][:6] + '_' + sheet + '_PrismIC50.csv'])
    else:
        csv_path = os.path.sep.join([folder_path, file.split('_')[0][:6] + '_' + sheet + '_Prism.csv'])
    csv_df = csv_df.iloc[::-1]
    csv_df.to_csv(csv_path, index = False)


    #This is for the second compound's csv file
    temp_list = []
    temp = new_df['10-fold'][30:60]/mean(new_df['10-fold'][63:66])*100
    temp = temp.reset_index()#Resets the index
    temp = temp.drop(['index'], axis =1)#Removes the messed up index column
    for i in range(0,30,3):
        temp_list.append((concentration_list[int(i/3)], temp['10-fold'][i], temp['10-fold'][i+1], temp['10-fold'][i+2]))

    csv_df = pd.DataFrame(temp_list, columns = ['Concentration', '% Control', 'x', 'y'])
    if sheet == '20-HETE':
        csv_path = os.path.sep.join([folder_path,'UPMP' + file.split('_')[0][6:9] + '_' + sheet + '_PrismIC50.csv'])
    else:
        csv_path = os.path.sep.join([folder_path,'UPMP' + file.split('_')[0][6:9] + '_' + sheet + '_Prism.csv'])
    csv_df = csv_df.iloc[::-1]
    csv_df.to_csv(csv_path, index = False)

    
def remove_files(folder_path):
    import shutil

    for file in os.listdir(folder_path):
        if file.split('_')[-1] == 'Prism.csv' or file.split('_')[-1] == 'PrismIC50.csv' :
            file_path = os.path.sep.join([folder_path, file])
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(e, 'This needs to be fixed.')
        else:
            pass

def main():
    folder_path = prompt() #Get the folder path
    new_df = open_file_IC50(folder_path) #Open the files, read them, and write the results file
    #prism_path = find_prism() #Finds the location of Prism
    #create_prism_script(folder_path) #Writes the Prism script
    #create_prism_script_regcurve(folder_path)
    #run_prism_script(prism_path, 'prism_script_regcurve.pzc') #Runs the Prism script for the other metabolites
    #run_prism_script(prism_path, 'prism_script.pzc')#Generates IC50
    remove_files(folder_path)#Removes the extra files that are not needed

    
main()
print('Done!')
