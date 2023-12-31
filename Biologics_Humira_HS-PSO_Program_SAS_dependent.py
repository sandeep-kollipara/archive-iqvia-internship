import os
import sys
import pandas
import numpy
import datetime
from multiprocessing import cpu_count, Pool
import win32api, win32process, win32con

'''
# 1: Read SAS file to VAR type 

def ReadSAStoVAR(sasfilename):
    sasvar=SAS7BDAT(sasfilename+".sas7bdat");
    return sasvar;

# 2: Write VAR type to TXT file (Obsolete)

def WriteVARtoTXT(sasvar):
    texter=open("VARtoTXT.txt", "w+");
    for row in sasvar:
        for word in row:
            if len(str(word)) > 7 and len(str(word)) < 16:
                texter.write("%s\t\t\t" % word);
            elif len(str(word)) >= 16 and len(str(word)) < 24 :
                texter.write("%s\t\t" % word);
            elif len(str(word)) >= 24:
                texter.write("%s\t" % word);
            else:
                texter.write("%s\t\t\t\t" % word);
        texter.write("\n");
    texter.close();

# 3: Write SAS file to TXT file (Obsolete)

def WriteSAStoTXT(sasfilename):
    texter=open("SAStoTXT.txt", "w+");
    with SAS7BDAT(sasfilename+".sas7bdat") as f:
        for row in f:
            for word in row:
                if len(str(word)) > 7 and len(str(word)) < 16:
                    texter.write("%s\t\t\t" % word);
                elif len(str(word)) >= 16 and len(str(word)) < 24 :
                    texter.write("%s\t\t" % word);
                elif len(str(word)) >= 24:
                    texter.write("%s\t" % word);
                else:
                    texter.write("%s\t\t\t\t" % word);
            texter.write("\n");
    texter.close();'''

# 4: Divide Transaction Rows by Units (Verified)

def DivideRowsbyUnits(sasframe):                        #The Units column doesn't update but the rows replicate.
    tempframe=sasframe;
    for index, row in sasframe.iterrows():              #Problem Encountered (Solved): Original Dataframe's Unit column's first element becomes 1. 
        if int(sasframe['Units'].loc[index]) > 1:
            units=int(sasframe['Units'][index]);
            units-=1;
            copyframe=tempframe.loc[[index]];
            copyframe['Units']=1;                       #Number of 'SettingWithCopyWarning's received.
            for i in range(int(units)):
                tempframe=pandas.concat([tempframe, copyframe]);
            tempframe['Units'].loc[[index]]=1;
            print("Multi unit encountered");
    return tempframe;

# 5: Convert SAS Date to Python Date
    
def DateSAStoPython(sasdate):
    sasdate=int(sasdate);
    leapyears=sasdate/1461;
    remyears=(sasdate%1461)/365;
    remdays=(sasdate%1461)%365;
    year=1960+4*int(leapyears)+int(remyears);
    if int(remyears)==3:                                #Leap Year
        if remdays <= 182:
            if remdays <= 91:
                if remdays <= 31:                       #January
                    month=1;
                    date=remdays;
                elif remdays > 31 and remdays <= 60:    #February
                    month=2;
                    date=remdays-31;
                else:                                   #March
                    month=3;
                    date=remdays-60;
            else:
                if remdays <= 121:                      #April
                    month=4;
                    date=remdays-91;
                elif remdays > 121 and remdays <= 152:  #May
                    month=5;
                    date=remdays-121;
                else:                                   #June
                    month=6;
                    date=remdays-152;
        else:
            if remdays <= 274:
                if remdays <= 213:                      #July
                    month=7;
                    date=remdays-182;
                elif remdays > 213 and remdays <= 244:  #August
                    month=8;
                    date=remdays-213;
                else:                                   #September
                    month=9;
                    date=remdays-244;
            else:
                if remdays <= 305:                      #October
                    month=10;
                    date=remdays-274;
                elif remdays > 305 and remdays <= 335:  #November
                    month=11;
                    date=remdays-305;
                else:                                   #December
                    month=12;
                    date=remdays-335;
    else:                                               #Normal Year
        if remdays <= 181:
            if remdays <= 90:
                if remdays <= 31:                       #January
                    month=1;
                    date=remdays;
                elif remdays > 31 and remdays <= 59:    #February
                    month=2;
                    date=remdays-31;
                else:                                   #March
                    month=3;
                    date=remdays-59;
            else:
                if remdays <= 120:                      #April
                    month=4;
                    date=remdays-90;
                elif remdays > 120 and remdays <= 151:  #May
                    month=5;
                    date=remdays-120;
                else:                                   #June
                    month=6;
                    date=remdays-151;
        else:
            if remdays <= 273:
                if remdays <= 212:                      #July
                    month=7;
                    date=remdays-181;
                elif remdays > 212 and remdays <= 243:  #August
                    month=8;
                    date=remdays-212;
                else:                                   #September
                    month=9;
                    date=remdays-243;
            else:
                if remdays <= 304:                      #October
                    month=10;
                    date=remdays-273;
                elif remdays > 304 and remdays <= 334:  #November
                    month=11;
                    date=remdays-304;
                else:                                   #December
                    month=12;
                    date=remdays-334;
    pydate=datetime.date(year, month, date);
    print(str(year)+" year and "+str(month)+" month and "+str(date)+" day.");
    return pydate;

# 6: Convert Python Date to SAS Date

def DatePythontoSAS(pydate):
    date=pydate.day;
    month=pydate.month;
    year=pydate.year;
    monthdays=[31,28,31,30,31,30,31,31,30,31,30,31];
    leapmonthdays=[31,29,31,30,31,30,31,31,30,31,30,31];
    daysofmonths=0;
    if int(year)%4!=0:
        for i in range(int(month)-1):
            daysofmonths+=monthdays[i];
    else:
        for i in range(int(month)-1):
            daysofmonths+=leapmonthdays[i];
    sasdate=(int(year)-1960)*365+int((int(year-1)-1960)/4)+int(daysofmonths)+int(date);
    return sasdate;

# 7: Input and Parse Python Date

def ReadandParseDateinPython():
    date_entry=input('Enter the date in YYYY-MM-DD format:');
    year,month,day=map(int, date_entry.split('-'));
    pydate=datetime.date(year,month,day);
    return pydate;

# 8: Check No. of patients in a Dataframe by 'P_ID':

def NumofPatients(dataframe):
    dfgroup=dataframe.groupby('P_ID');
    i=0;
    for patient_id, patient_stats in dfgroup:
        i+=1;
    return i;

# 9: Check for Patient No. in the dataframe:

def IsPatientNumin(dataframe, patient_id):
    answer=False;
    for p_id in dataframe['P_ID']:
        if int(p_id) == int(patient_id):
            answer=True;
    return answer;

# 10: Isolate Patient data from a dataframe

def RetrievePatientfrom(dataframe, p_id):
    dfgroup=dataframe.groupby('P_ID');
    for patient_id, patient_stats in dfgroup:
        if patient_id == p_id:
            patientframe=patient_stats;
    return patientframe;

# 11: Get patient IDs from a data frame

def RetrievePatientIDsfrom(dataframe):
    patientframe={
        'P_ID':[],
    }
    pframe=pandas.DataFrame(patientframe);
    dfgroup=dataframe.groupby('P_ID');
    for patient_id, patient_stats in dfgroup:
        pframe=pframe.append({'P_ID':patient_id}, ignore_index=True);
    return pframe;

# 12: Multiprocessing Pack Numbering

#def PackNumbering(packframe, patient_stats):

# 13: Multiprocessing Therapy Duration

#def AssignRXDuration(rx):

def AssignPatientLevelSpec(threadframe):
    HighPriority();
    emptyframe2=pandas.DataFrame(columns=['combomol', 'DDMS_PHA', 'TransactionDate', 'FCC', 'Units', 'ShortCode','P_ID', 'TransactionMonth', 'atc', 'prod', 'pack', 'str_unit','str_meas', 'gene', 'DoctorSpeciality', 'CU', 'PSIZE', 'numdax', 'DoctorClass', 'PatientLevelClass']);
    emptyframe=pandas.DataFrame(columns=['combomol', 'DDMS_PHA', 'TransactionDate', 'FCC', 'Units', 'ShortCode','P_ID', 'TransactionMonth', 'atc', 'prod', 'pack', 'str_unit','str_meas', 'gene', 'DoctorSpeciality', 'CU', 'PSIZE', 'numdax', 'DoctorClass']);
    rframe=dframe=uframe=emptyframe2;
    oframe=emptyframe;
    patientframe={
            'PatientNum':[],
            'PatientDocClass':[]
    }
    pframe=pandas.DataFrame(patientframe);
    excelconst=pandas.read_excel('Biologics_Humira_HS-PSO_Program_Settings.xlsx', sheet_name='SpecialityBiologics');
    tfgroup=threadframe.groupby('P_ID');
    for patient_id, patient_stats in tfgroup:
        rpositive=dpositive=False;                      #@gpositive var removed
        for molecule in patient_stats['combomol']:
            for const in excelconst['Rheumato_Molecules']:
                if molecule==const:
                    rpositive=True;
            for const in excelconst['Dermato_Molecules']:
                if molecule==const:
                    dpositive=True;
            '''for const in excelconst['Gastro_Molecules']:    #not_implemented
                if molecule==const:
                    gpositive=True;'''
        if rpositive==True and dpositive==False:
            pframe=pframe.append({'PatientNum':patient_id,'PatientDocClass':'Rheumato'}, ignore_index=True);
            tempframe=patient_stats;
            tempframe=tempframe.copy();                      #bad_practice
            tempframe['PatientLevelClass']='Rheumato';
            rframe=pandas.concat([rframe, tempframe]);
        elif rpositive==False and dpositive==True:
            pframe=pframe.append({'PatientNum':patient_id,'PatientDocClass':'Dermato'}, ignore_index=True);
            tempframe=patient_stats;
            tempframe=tempframe.copy();                      #bad_practice
            tempframe['PatientLevelClass']='Dermato';
            dframe=pandas.concat([dframe, tempframe]);
        elif rpositive==True and dpositive==True:
            pframe=pframe.append({'PatientNum':patient_id,'PatientDocClass':'Unknown'}, ignore_index=True);
            tempframe=patient_stats;
            tempframe=tempframe.copy();                      #bad_practice
            tempframe['PatientLevelClass']='Unknown';
            uframe=pandas.concat([uframe, tempframe]);
        else:
            pframe=pframe.append({'PatientNum':patient_id,'PatientDocClass':'Others'}, ignore_index=True);
            tempframe=patient_stats;
            tempframe=tempframe.copy();                      #bad_practice
            #tempframe['PatientLevelClass']='Others';
            oframe=pandas.concat([oframe, tempframe]);
        print("Patient No."+str(patient_id)+" processed.");
    threadframelist=[pframe, rframe, dframe, uframe, oframe];
    return threadframelist;

def AssignPatientLevelSpectoExceptions(threadframe):
    HighPriority();
    emptyframe2=pandas.DataFrame(columns=['combomol', 'DDMS_PHA', 'TransactionDate', 'FCC', 'Units', 'ShortCode','P_ID', 'TransactionMonth', 'atc', 'prod', 'pack', 'str_unit','str_meas', 'gene', 'DoctorSpeciality', 'CU', 'PSIZE', 'numdax', 'DoctorClass', 'PatientLevelClass']);
    gframe=eframe=rframe=dframe=emptyframe2;
    tfgroup=threadframe.groupby('P_ID');
    for patient_id, patient_stats in tfgroup:
        g=r=d=irr=0;
        for spec in patient_stats['DoctorClass']:
            if spec == 'Gastro':
                g+=1;
            elif spec == 'Rheumato':
                r+=1;
            elif spec == 'Dermato':
                d+=1;
            elif spec == 'Irrelevant':
                irr+=1;
            else:
                continue;
        if g > r and g > d and g >= irr:
            tempframe=patient_stats;
            tempframe=tempframe.copy();                      #bad_practice
            tempframe['PatientLevelClass']='Gastro';
            gframe=pandas.concat([gframe, tempframe]);
        elif r > g and r > d and r >= irr:
            tempframe=patient_stats;
            tempframe=tempframe.copy();                      #bad_practice
            tempframe['PatientLevelClass']='Rheumato';
            rframe=pandas.concat([rframe, tempframe]);
        elif d > g and d > r and d >= irr:
            tempframe=patient_stats;
            tempframe=tempframe.copy();                      #bad_practice
            tempframe['PatientLevelClass']='Dermato';
            dframe=pandas.concat([dframe, tempframe]);
        elif r==g and r > d and r >= irr:
            tempframe=patient_stats;
            tempframe=tempframe.copy();                      #bad_practice
            tempframe['PatientLevelClass']='Rheumato';
            rframe=pandas.concat([rframe, tempframe]);
        elif r==d and d > g and d >= irr:
            tempframe=patient_stats;
            tempframe=tempframe.copy();                      #bad_practice
            tempframe['PatientLevelClass']='Rheumato';
            rframe=pandas.concat([rframe, tempframe]);
        elif g==d and g > r and g >= irr:
            tempframe=patient_stats;
            tempframe=tempframe.copy();                      #bad_practice
            tempframe['PatientLevelClass']='Gastro';
            gframe=pandas.concat([gframe, tempframe]);
        elif r==g and g==d and d >= irr:
            tempframe=patient_stats;
            tempframe=tempframe.copy();                      #bad_practice
            tempframe['PatientLevelClass']='Rheumato';
            rframe=pandas.concat([rframe, tempframe]);
        elif irr==g and g > r and g > d:
            tempframe=patient_stats;
            tempframe=tempframe.copy();                      #bad_practice
            tempframe['PatientLevelClass']='Gastro';
            gframe=pandas.concat([gframe, tempframe]);
        elif irr==r and r > g and r > d:
            tempframe=patient_stats;
            tempframe=tempframe.copy();                      #bad_practice
            tempframe['PatientLevelClass']='Rheumato';
            rframe=pandas.concat([rframe, tempframe]);
        elif irr==d and d > r and d > g:
            tempframe=patient_stats;
            tempframe=tempframe.copy();                      #bad_practice
            tempframe['PatientLevelClass']='Dermato';
            dframe=pandas.concat([dframe, tempframe]);
        elif irr==g and g==r and r==d:
            tempframe=patient_stats;
            tempframe=tempframe.copy();                      #bad_practice
            tempframe['PatientLevelClass']='Rheumato';
            rframe=pandas.concat([rframe, tempframe]);
        else:
            tempframe=patient_stats;
            tempframe=tempframe.copy();                      #bad_practice
            tempframe['PatientLevelClass']='Error';
            eframe=pandas.concat([eframe, tempframe]);
        print("Exception Patient No."+str(patient_id)+" processed.");
    threadframelist=[gframe, rframe, dframe, eframe];
    return threadframelist;

def PackNumbering(threadframe):
    HighPriority();
    print("Multithreading starts...", flush=0);
    i=0;
    global packnumbermatrix;
    for index, rx in threadframe.iterrows():
        if i==0:
            threadframe['PackNumber'].iloc[[i]]=1;
            j=1;
        else:
            if threadframe['P_ID'].iloc[[i-1]].item() == threadframe['P_ID'].iloc[[i]].item() \
            and threadframe['combomol'].iloc[[i-1]].item() == threadframe['combomol'].iloc[[i]].item():
                j+=1;
            elif threadframe['P_ID'].iloc[[i-1]].item() == threadframe['P_ID'].iloc[[i]].item() \
            and threadframe['combomol'].iloc[[i-1]].item() != threadframe['combomol'].iloc[[i]].item():
                j=1;
            else:
                j=1;
            threadframe['PackNumber'].iloc[[i]]=j;
        print(str(i+1)+" RXs indexed.",flush=True);
        i+=1;
    return threadframe;
    print("...end of Multithreading", flush=0);

def PartitionDataFramebyPatientIDs(refframe):
    rflength=len(refframe);
    midindex=int(rflength/2);
    lastpatient=False;
    while not lastpatient:
        if refframe['P_ID'].iloc[[midindex+1]].item() == refframe['P_ID'].iloc[[midindex]].item() \
        and midindex + 2 < rflength:
            midindex+=1;
        else:
            lastpatient=True;
    print("Partition done.")
    partframe1=refframe[:midindex+1];
    partframe2=refframe[midindex+1:];
    partslist=[partframe1, partframe2];
    return partslist;

def PartitionDataFramebyCores(refframe, parts):
    if parts==1:
        partslist=[refframe];
        print("No partioning as only 1 core.");
        return partslist;
    elif parts==2:
        partslist=PartitionDataFramebyPatientIDs(refframe);
        print("Double partition done");
        return partslist;
    elif parts==4:
        partslist=PartitionDataFramebyPatientIDs(refframe);
        partframe1=partslist[0];
        partframe2=partslist[1];
        partslist1=PartitionDataFramebyPatientIDs(partframe1);
        partslist2=PartitionDataFramebyPatientIDs(partframe2);
        partslist=partslist1+partslist2;
        print("Quadruple partition done");
        return partslist;
    else:
        partslist=PartitionDataFramebyPatientIDs(refframe);
        partframe1=partslist[0];
        partframe2=partslist[1];
        partslist1=PartitionDataFramebyPatientIDs(partframe1);
        partslist2=PartitionDataFramebyPatientIDs(partframe2);
        partframe1=partslist1[0];
        partframe2=partslist1[1];
        partframe3=partslist2[0];
        partframe4=partslist2[1];
        partslist1=PartitionDataFramebyPatientIDs(partframe1);
        partslist2=PartitionDataFramebyPatientIDs(partframe2);
        partslist3=PartitionDataFramebyPatientIDs(partframe3);
        partslist4=PartitionDataFramebyPatientIDs(partframe4);
        partslist=partslist1+partslist2+partslist3+partslist4;
        print("Octuple partition done");
        return partslist;

def HighPriority():
    """ Set the priority of the process to high (Source:Stackexchange) """
    try:
        sys.getwindowsversion();
    except AttributeError:
        isWindows = False;
    else:
        isWindows = True;
    if isWindows:
        # Based on:
        #   "Recipe 496767: Set Process Priority In Windows" on ActiveState
        #   http://code.activestate.com/recipes/496767/
        pid = win32api.GetCurrentProcessId();
        handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, True, pid);
        win32process.SetPriorityClass(handle, win32process.NORMAL_PRIORITY_CLASS);
    else:
        os.nice(10);

if __name__=='__main__':
    
    # --------------------------------PHASE I-------------------------------------
    
    print("Importing Biologics SAS dataset...");
    filelocs=pandas.read_excel("Biologics_Humira_HS-PSO_Program_Settings.xlsx", sheet_name='FileLocations');
    importedsasfile=filelocs['Address'][0];
    sasframe=pandas.read_sas(importedsasfile, format='sas7bdat', encoding='iso-8859-1');
    sasframe['DoctorClass']=sasframe['DoctorSpeciality'];
    sasframe['DoctorClass'].replace(['31','32','47'],['Dermato','Gastro','Rheumato'], inplace=True);
    for num in range(sasframe['DoctorClass'].size):
        if sasframe['DoctorClass'][num]=='':
            sasframe['DoctorClass'].at[num]='Irrelevant';
        elif sasframe['DoctorClass'][num]!='Rheumato' and sasframe['DoctorClass'][num]!='Gastro' and sasframe['DoctorClass'][num]!='Dermato':
            sasframe['DoctorClass'].at[num]='Irrelevant';
    print(sasframe);
    sasframe=DivideRowsbyUnits(sasframe);
    
    # ---------------------------------PHASE II-----------------------------------
    
    #sasframe['P_LevelClass']='';
    emptyframe2=pandas.DataFrame(columns=['combomol', 'DDMS_PHA', 'TransactionDate', 'FCC', 'Units', 'ShortCode','P_ID', 'TransactionMonth', 'atc', 'prod', 'pack', 'str_unit','str_meas', 'gene', 'DoctorSpeciality', 'CU', 'PSIZE', 'numdax', 'DoctorClass', 'PatientLevelClass']);
    emptyframe=pandas.DataFrame(columns=['combomol', 'DDMS_PHA', 'TransactionDate', 'FCC', 'Units', 'ShortCode','P_ID', 'TransactionMonth', 'atc', 'prod', 'pack', 'str_unit','str_meas', 'gene', 'DoctorSpeciality', 'CU', 'PSIZE', 'numdax', 'DoctorClass']);
    rframe=dframe=uframe=emptyframe2;
    oframe=emptyframe;
    patientframe={
            'PatientNum':[],
            'PatientDocClass':[]
    }
    pframe=pandas.DataFrame(patientframe);
    if cpu_count() < 2:
        partitions=1;
    elif cpu_count() < 4:
        partitions=2;
    elif cpu_count() < 8:
        partitions=4;
    else:
        partitions=8;
    sasframe=sasframe.sort_values(by=['P_ID', 'numdax', 'combomol'], ascending=[True, True, True]);
    print("Dataframe sorted by Patient ID, Transaction date and Molecule.");
    sasframe['PatientLevelClass']='';
    sasflength=len(sasframe);
    processedpartitionframelist=[];
    partitionframelist=PartitionDataFramebyCores(sasframe, partitions);
    pool=Pool();
    processedpartitionframelist.append(pool.map(AssignPatientLevelSpec, partitionframelist));
    pool.close();
    pool.join();
    for i in range(partitions):
        pframe=pandas.concat([pframe, processedpartitionframelist[0][i][0]]);
        rframe=pandas.concat([rframe, processedpartitionframelist[0][i][1]]);
        dframe=pandas.concat([dframe, processedpartitionframelist[0][i][2]]);
        uframe=pandas.concat([uframe, processedpartitionframelist[0][i][3]]);
        oframe=pandas.concat([oframe, processedpartitionframelist[0][i][4]]);
    print("Phase II: Part I execution ended at"+str(datetime.datetime.now().time()));
    
    # ------------------------PHASE II: Exception Handling------------------------
    
    emptyframe2=pandas.DataFrame(columns=['combomol', 'DDMS_PHA', 'TransactionDate', 'FCC', 'Units', 'ShortCode','P_ID', 'TransactionMonth', 'atc', 'prod', 'pack', 'str_unit','str_meas', 'gene', 'DoctorSpeciality', 'CU', 'PSIZE', 'numdax', 'DoctorClass', 'PatientLevelClass']);
    gframe=eframe=emptyframe2;
    oframe=oframe.sort_values(by=['P_ID', 'numdax', 'combomol'], ascending=[True, True, True]);
    print("Dataframe sorted by Patient ID, Transaction date and Molecule.");
    oflength=len(oframe);
    processedpartitionframelist=[]
    partitionframelist=PartitionDataFramebyCores(oframe, partitions);
    pool=Pool();
    processedpartitionframelist.append(pool.map(AssignPatientLevelSpectoExceptions, partitionframelist));
    pool.close();
    pool.join();
    for i in range(partitions):
        gframe=pandas.concat([gframe, processedpartitionframelist[0][i][0]]);
        rframe=pandas.concat([rframe, processedpartitionframelist[0][i][1]]);
        dframe=pandas.concat([dframe, processedpartitionframelist[0][i][2]]);
        eframe=pandas.concat([eframe, processedpartitionframelist[0][i][3]]);
    print("Phase II: Part II execution ended at"+str(datetime.datetime.now().time()));
    
    # ----------------- PHASE III: i. Flag Product Transactions ------------------
    
    emptyframe3=pandas.DataFrame(columns=['combomol', 'DDMS_PHA', \
                                          'TransactionDate', 'FCC', 'Units', \
                                          'ShortCode','P_ID', 'TransactionMonth', \
                                          'atc', 'prod', 'pack', 'str_unit', \
                                          'str_meas', 'gene', 'DoctorSpeciality', \
                                          'CU', 'PSIZE', 'numdax', 'DoctorClass', \
                                          'PatientLevelClass', 'ProdTransaction']);
    npframe=emptyframe3;
    dfgroup=dframe.groupby('P_ID');
    for patient_id, patient_stats in dfgroup:
        tempframe=patient_stats;
        tempframe=tempframe.sort_values(by=['numdax', 'combomol'], ascending=[True, True]);
        tempframe['prevnumdax']=tempframe['numdax'].shift(1);
        tempframe['prevnumdax'].iloc[[0]]=0;
        tempframe['interval']=tempframe.apply(lambda rx: rx['numdax']-rx['prevnumdax'], axis=1);
        tempframe['ProdTransaction']=tempframe['interval'].apply(lambda rx: 'New' if rx >= 540 else '');
        '''
        for index, rx in tempframe.iterrows():
            if int(rx['interval']) >= 540:
                rx['ProdTransaction']=='New';
                tempframe['firstnumdax']=rx['numdax'];
            else:
                rx['ProdTransaction']=='';
        '''
        tempframe.drop(['prevnumdax', 'interval'], axis=1);
        npframe=pandas.concat([npframe, tempframe]);
        print('Patient No.'+str(patient_id)+'\'s transactions flagged.');
    
        '''
        i=0;
        newi=0;
        for sasdate in tempframe['numdax']:
            currprod=tempframe['combomol'].iloc[[i]];
            if i==0:
                tempframe['ProdTransaction'].iloc[[i]]='New';
                prevprod=currprod;
                prevdate=sasdate;
            else:
                if int(prevdate + 540) > int(sasdate):
                    if prevprod.item() == currprod.item():
                        tempframe['ProdTransaction'].iloc[[i]]='Repeat';
                    else:
                        tempframe['ProdTransaction'].iloc[[i]]='Switch';
                else:  
                    tempframe['ProdTransaction'].iloc[[i]]='New';
                prevprod=currprod;
                prevdate=sasdate;
            i+=1;
        npframe=pandas.concat([npframe, tempframe]);
        print('Patient No.'+str(patient_id)+'\'s transactions flagged.');
        '''
    
    # ----------------- PHASE III: ii. Filter New Humira Patients ----------------
    
    npfgroup=npframe.groupby('P_ID');
    newhumirafilterframe=nonhumirafilterframe=emptyframe3;
    for patient_id, patient_stats in npfgroup:
        humirauser=False;
        tempframe=patient_stats;
        i=0;
        for molecule in tempframe['combomol']:
            if 'ADALIMUMAB' in str(molecule) \
            and tempframe['ProdTransaction'].iloc[[i]].item()=='New':
                humirauser=True;
            i+=1;
        i=0;
        for molecule in tempframe['combomol']:
            if "SER.PREREMPL 20MG ENF 2 .2ML" in tempframe['pack'].iloc[[i]].item() \
            or "SOL.INJ. 40MG ENF 2 .8ML" in  tempframe['pack'].iloc[[i]].item() \
            or "TECFIDERA" in tempframe['prod'].iloc[[i]].item() \
            or "MABTHERA" in tempframe['prod'].iloc[[i]].item():
                humirauser=False;
            i+=1;
        if humirauser:
            print('New Humira user found.');
            newhumirafilterframe=pandas.concat([newhumirafilterframe, tempframe]);
        else:
            nonhumirafilterframe=pandas.concat([nonhumirafilterframe, tempframe]);
    
    # ------------------ PHASE III: iii. HS/PsO Indication Split -----------------
    
    newhumirafilterframe['Flag_Till_Now']=numpy.nan;
    emptyframe6=pandas.DataFrame(columns=['combomol', 'DDMS_PHA', \
                                          'TransactionDate', 'FCC', 'Units', \
                                          'ShortCode','P_ID', 'TransactionMonth', \
                                          'atc', 'prod', 'pack', 'str_unit', \
                                          'str_meas', 'gene', 'DoctorSpeciality', \
                                          'CU', 'PSIZE', 'numdax', 'DoctorClass', \
                                          'PatientLevelClass', 'ProdTransaction', \
                                          'Flag_Till_Now']);
    
    # ----------------- PHASE III: iii. a.) September 2016 Update ----------------
    
    hsp1frame=oldhpframe=emptyframe6;
    nhffgroup=newhumirafilterframe.groupby('P_ID');
    for patient_id, patient_stats in nhffgroup:
        newonhumirapostsept2016=False;
        tempframe=patient_stats;
        i=0;
        for buydate in tempframe['TransactionDate']:
            if tempframe['combomol'].iloc[[i]].item()=='ADALIMUMAB' \
            and tempframe['ProdTransaction'].iloc[[i]].item()=='New' \
            and int(buydate) >= 20160901:
                newonhumirapostsept2016=True;
            elif tempframe['combomol'].iloc[[i]].item()=='ADALIMUMAB' \
            and tempframe['ProdTransaction'].iloc[[i]].item()=='New' \
            and int(buydate) < 20160901:
                tempframe['ProdTransaction'].iloc[[i]].item()=='Old'
            i+=1;
        if newonhumirapostsept2016==True:
            hsp1frame=pandas.concat([hsp1frame, tempframe]);
        else:
            oldhpframe=pandas.concat([oldhpframe, tempframe]);
    
    # ------------- PHASE III: iii. b) Anti-TNF to PsO Classification ------------
    
    print("Initializing Anti-TNF to PSO classification.");
    hsp2frame=pso1frame=emptyframe6;
    hsp1fgroup=hsp1frame.groupby('P_ID');
    for patient_id, patient_stats in hsp1fgroup:
        onlyhumira=True;
        tempframe=patient_stats;
        i=0;
        for molecule in tempframe['combomol']:
            if 'ADALIMUMAB' not in str(molecule):
                onlyhumira=False;
            i+=1;
        if onlyhumira:
            print("HS probable patient No."+str(patient_id)+" detected.");
            hsp2frame=pandas.concat([hsp2frame, tempframe]);
        else:
            print("PsO Patient No."+str(patient_id)+" detected.");
            pso1frame=pandas.concat([pso1frame, tempframe]);
    print("Concluding Anti-TNF to PSO classification.");
    
    # --------------- PHASE III: iii. c.) 1st Month Injections Basis -------------
    
    print("Initializing Monthly injection based classification.");
    hsp3frame=pso2frame=hs1frame=emptyframe6;
    hsp2fgroup=hsp2frame.groupby('P_ID');
    for patient_id, patient_stats in hsp2fgroup:
        sumofunits=0;
        tempframe=patient_stats;
        tempframe=tempframe.sort_values(by=['numdax'], ascending=True);
        i=0;
        for buydate in tempframe['numdax']:
            if tempframe['ProdTransaction'].iloc[[i]].item() == 'New':
                firstdate=buydate;
                #break;
            i+=1;
        i=0;
        for buydate in tempframe['numdax']:
            if 0 <= int(buydate - firstdate) <= 30:
                #sumofunits+=int(tempframe['Units'].iloc[[i]].item());
                sumofunits+=int(int(tempframe['PSIZE'].iloc[[i]].item())*int(tempframe['str_unit'].iloc[[i]].item())*int(tempframe['Units'].iloc[[i]].item()))/80;
            i+=1;
        if int(sumofunits) >= 3:
            print("HS patient No."+str(patient_id)+" detected.");
            hs1frame=pandas.concat([hs1frame, tempframe]);
        elif int(sumofunits) == 2:
            print("HS probable patient No."+str(patient_id)+" detected.");
            hsp3frame=pandas.concat([hsp3frame, tempframe]);
        else:
            print("PsO patient No."+str(patient_id)+" detected.");
            pso2frame=pandas.concat([pso2frame, tempframe]);
    print("Concluding Monthly injection based classification.");
    
    # --------------------- PHASE III: iii. d.) Server Access --------------------
    
    #created hsp3masterframe including comedications for patients in hs3frame and sort by P_ID and TransactionDate/numdax
    patproframe=RetrievePatientIDsfrom(hsp1frame);
    patproframe.to_excel("PatientProfilePID.xlsx", sheet_name='Patient ID', index=False);
    print("Exporting the patients to import master data.");
    exportpatientframe=RetrievePatientIDsfrom(hsp3frame);
    exportpatientframe.to_excel("ComedicationID.xlsx", sheet_name="Patient ID", index=False);
    print("Exporting the patients to import patient profile data.");
    '''
    hsp3masterframe['Batch_ID']=numpy.nan;
    hsp3masterframe['Sequence']=numpy.nan;
    hsp3masterframe['BasketNumber']=numpy.nan;
    hsp3masterframe['RX_ID']=numpy.nan;
    hsp3masterframe['TransactionType']=numpy.nan;
    hsp3masterframe['P_Gender']=numpy.nan;
    hsp3masterframe['P_Loyalty']=numpy.nan;
    hsp3masterframe['DCI_VOS']=numpy.nan;
    '''
    input("Please place the master SAS dataset (name=master1) and patient profile SAS dataset (name=patprofil) in the program's folder and press ENTER to continue.");
    print("Importing Master comedication dataset.");
    #importedsasfile2=input("Enter the filename of the Master SAS dataset to be imported(without ext.):")
    importedsasfile2=filelocs['Address'][1];
    penultimateframe=pandas.read_sas(importedsasfile2, format='sas7bdat', encoding='iso-8859-1');
    penultimateframe.reset_index(drop=True);
    if 'COMBOMOL' in penultimateframe.columns:
        penultimateframe['combomol']=penultimateframe['COMBOMOL'];
    if 'PROD' in penultimateframe.columns:
        penultimateframe['prod']=penultimateframe['PROD'];
    if 'PSIZE' in penultimateframe.columns:
        penultimateframe['psize']=penultimateframe['PSIZE'];
    if 'PACK' in penultimateframe.columns:
        penultimateframe['pack']=penultimateframe['PACK'];
    pufgroup=penultimateframe.groupby('P_ID');
    masterframe=penultimateframe[0:0];
    for patient_id_10, patient_stats_10 in pufgroup:
        exception=False;
        tempframe=patient_stats_10;
        i=0;
        for molecule in tempframe['combomol']:
            if "SER.PREREMPL 20MG ENF 2 .2ML" in str(tempframe['pack'].iloc[[i]].item()) \
            or "SOL.INJ. 40MG ENF 2 .8ML" in  str(tempframe['pack'].iloc[[i]].item()) \
            or "TECFIDERA" in str(tempframe['prod'].iloc[[i]].item()) \
            or "MABTHERA" in str(tempframe['prod'].iloc[[i]].item()):
                exception=True;
            i+=1;
        if exception:
            continue;
        else:
            masterframe=pandas.concat([masterframe, tempframe]);
    for index, row in masterframe.iterrows():
        if str(masterframe['combomol'].loc[index])=='ADALIMUMAB':
            masterframe=masterframe.drop(index=[index]);
    masterframe2=masterframe[0:0];
    masterfgroup=masterframe.groupby('P_ID');
    for patient_id_ex in exportpatientframe['P_ID']:
        for patient_id_master, tempframe_master in masterfgroup:
            if patient_id_ex == patient_id_master:
                masterframe2=pandas.concat([masterframe2, tempframe_master]);
    hsp3masterframe=pandas.concat([masterframe2, hsp3frame], sort=True);
    
    patproframe=RetrievePatientIDsfrom(hsp1frame);
    patproframe.to_excel("PatProFrame.xlsx", sheet_name='PatProfil', index=False);
    
    '''
    
    hsp3masterframe=pandas.merge(hsp3frame, masterframe, \
                                 how='outer', \
                                 left_on=['combomol', 'DDMS_PHA', \
                                          'TransactionDate', 'FCC', 'Units', \
                                          'ShortCode', 'P_ID', 'prod', 'pack', \
                                          'PSIZE'], \
                                 right_on=['combomol', 'DDMS_PHA', \
                                           'TransactionDate', 'FCC', 'Units', \
                                           'ShortCode', 'P_ID', 'prod', 'pack', \
                                           'psize']);
    
    mfgroup=masterframe.groupby('P_ID');
    hsp3fgroup=hsp3frame.groupby('P_ID');
    for patient_id_1, patient_stats_1 in hsp3fgroup:
        for patient_id_2, patient_stats_2 in mfgroup:
            if patient_id_1 == patient_id_2:
                tempframe=patient_stats_2;
                hsp3masterframe=pandas.concat([hsp3masterframe, tempframe]);
            else:
                continue;
    
    print("Combined Master comedication dataset with the original.");
    '''
    # - PHASE III: iii. e.) Antibiotic:Tetracycline & Rifampicin + Clindampicine -
    
    print("Initializing Antibiotic check on comedication.");
    emptyframe7=pandas.DataFrame(columns=['Batch_ID', 'Sequence', 'BasketNumber', \
                                          'RX_ID', 'TransactionType', 'P_Gender', \
                                          'P_Loyalty', 'DCI_VOS', 'combomol', \
                                          'DDMS_PHA', 'TransactionDate', 'FCC', \
                                          'Units', 'ShortCode','P_ID', \
                                          'TransactionMonth', 'atc', 'prod', \
                                          'pack', 'str_unit','str_meas', 'gene', \
                                          'DoctorSpeciality', 'CU', 'PSIZE', \
                                          'psize', 'numdax', 'DoctorClass', \
                                          'PatientLevelClass', 'ProdTransaction', \
                                          'Flag_Till_Now']);
    hsp4frame=hs2frame=emptyframe7;
    hsp3mfgroup=hsp3masterframe.groupby('P_ID');
    for patient_id, patient_stats in hsp3mfgroup:
        antibiotic1=antibiotic2=pre_clindamycin=pre_rifampicin=False;
        combidate=-1;
        tempframe=patient_stats;
        tempframe=tempframe.sort_values(by=['TransactionDate'], ascending=True);
        i=0;
        for molecule in tempframe['combomol']:
            if 'ADALIMUMAB' in str(molecule) \
            and tempframe['ProdTransaction'].iloc[[i]].item() == 'New':
                firstdate=tempframe['TransactionDate'].iloc[[i]].item();
            i+=1;
        i=0;
        for buydate in tempframe['TransactionDate']:
            if 'TETRACYCLINE' in str(tempframe['prod'].iloc[[i]].item()) \
            or 'TETRACYCLINE' in str(tempframe['combomol'].iloc[[i]].item()) \
            or 'LYMECYCLINE' in str(tempframe['prod'].iloc[[i]].item()) \
            or 'LYMECYCLINE' in str(tempframe['combomol'].iloc[[i]].item()) \
            or 'MINOCYCLINE' in str(tempframe['prod'].iloc[[i]].item()) \
            or 'MINOCYCLINE' in str(tempframe['combomol'].iloc[[i]].item()) \
            or 'DOXYCYCLINE' in str(tempframe['prod'].iloc[[i]].item()) \
            or 'DOXYCYCLINE' in str(tempframe['combomol'].iloc[[i]].item()):
                if True: #buydate < firstdate:
                    antibiotic1=True;
            #if 'RIFAMPICIN' in str(tempframe['combomol'].iloc[[i]].item()) \
            #and buydate < firstdate:
            if 'RIFAMPICIN' in str(tempframe['combomol'].iloc[[i]].item()):
                transpydate=datetime.date(int(buydate/10000), \
                                          int((buydate%10000)/100), \
                                          int(buydate%100));
                sasdate=DatePythontoSAS(transpydate);
                rxd=tempframe['psize'].iloc[[i]].item()*tempframe['Units'].iloc[[i]].item();
                if pre_clindamycin and combidate > sasdate:
                    antibiotic2=True;
                    break;
                elif pre_rifampicin and combidate > sasdate:
                    combidate=combidate+rxd;
                else:
                    combidate=sasdate+rxd;
                pre_rifampicin=True;
                pre_clindamycin=False;
            #if 'CLINDAMYCIN' in str(tempframe['combomol'].iloc[[i]].item()) \
            #and buydate < firstdate:
            if 'CLINDAMYCIN' in str(tempframe['combomol'].iloc[[i]].item()):
                transpydate=datetime.date(int(buydate/10000), \
                                          int((buydate%10000)/100), \
                                          int(buydate%100));
                sasdate=DatePythontoSAS(transpydate);
                rxd=tempframe['psize'].iloc[[i]].item()*tempframe['Units'].iloc[[i]].item();
                if pre_clindamycin and combidate > sasdate:
                    combidate=combidate+rxd;
                elif pre_rifampicin and combidate > sasdate:
                    antibiotic2=True;
                    break;
                else:
                    combidate=sasdate+rxd;
                pre_clindamycin=True;
                pre_rifampicin=False;
            i+=1;
        if antibiotic1 == True:
            print("HS patient No."+str(patient_id)+" with Tetracyclins detected.");
            hs2frame=pandas.concat([hs2frame, tempframe]);
        elif antibiotic2 == True:
            print("HS patient No."+str(patient_id)+" with CLINDAMYCIN + RIFAMPICIN detected.");
            hs2frame=pandas.concat([hs2frame, tempframe]);
        else:
            print("HS probable patient No."+str(patient_id)+" detected.");
            hsp4frame=pandas.concat([hsp4frame, tempframe]);
    print("Concluding Antibiotic check on comedication.");
    
    # -------------- PHASE III: iii. f.) Methotrexate & Cyclosporine -------------
    
    print("Initiializing Methotrexate & Cyclosporine check on comedication.");
    hsp5frame=pso3frame=emptyframe7;
    hsp4fgroup=hsp4frame.groupby('P_ID');
    for patient_id, patient_stats in hsp4fgroup:
        psopositive=False;
        tempframe=patient_stats;
        tempframe=tempframe.sort_values(by=['TransactionDate'], ascending=True);
        i=0;
        for molecule in tempframe['combomol']:
            if 'ADALIMUMAB' in str(molecule) \
            and tempframe['ProdTransaction'].iloc[[i]].item() == 'New':
                firstdate=tempframe['TransactionDate'].iloc[[i]].item();
            i+=1;
        i=0;
        for buydate in tempframe['TransactionDate']:
            #if 'METHOTREXATE' in str(tempframe['combomol'].iloc[[i]].item()) \
            #and buydate < firstdate:
            if 'METHOTREXATE' in str(tempframe['combomol'].iloc[[i]].item()):
                psopositive=True;
            #if 'CICLOSPORIN' in str(tempframe['combomol'].iloc[[i]].item()) \
            #and buydate < firstdate:
            if 'CICLOSPORIN' in str(tempframe['combomol'].iloc[[i]].item()):
                psopositive=True;
            i+=1;
        if psopositive:
            print("PsO patient No."+str(patient_id)+" detected.");
            pso3frame=pandas.concat([pso3frame, tempframe]);
        else:
            print("HS probable patient No."+str(patient_id)+" detected.");
            hsp5frame=pandas.concat([hsp5frame, tempframe]);
    print("Concluding Methotrexate & Cyclosporine check on comedication.");
    
    # ---------------------- PHASE III: iii. g.) Isobetadine --------------------
    
    print("Initializing Iso-betadine check on comedication.");
    hsp6frame=hs3frame=emptyframe7;
    hsp5fgroup=hsp5frame.groupby('P_ID');
    for patient_id, patient_stats in hsp5fgroup:
        hspositive=False;
        tempframe=patient_stats;
        tempframe=tempframe.sort_values(by=['TransactionDate'], ascending=True);
        i=0;
        for molecule in tempframe['combomol']:
            if 'ADALIMUMAB' in str(molecule) \
            and tempframe['ProdTransaction'].iloc[[i]].item() == 'New':
                firstdate=tempframe['TransactionDate'].iloc[[i]].item();
            i+=1;
        i=0;
        for buydate in tempframe['TransactionDate']:
            if 'ISO-BETADINE' in str(tempframe['combomol'].iloc[[i]].item()) \
            or 'ISO-BETADINE' in str(tempframe['prod'].iloc[[i]].item()):
                if buydate >= firstdate:
                    hspositive=True;
            i+=1;
        if hspositive:
            print("HS patient No."+str(patient_id)+" detected.");
            hs3frame=pandas.concat([hs3frame, tempframe]);
        else:
            print("HS probable patient No."+str(patient_id)+" detected.");
            hsp6frame=pandas.concat([hsp6frame, tempframe]);
    print("Concluding Iso-betadiene check on comedication.");
    
    # ----------------- PHASE III: iii. h.) Psoriasis Only Drugs -----------------
    
    print("Initializing Psoriasis-only drug check on comedication.");
    hsp7frame=pso4frame=emptyframe7;
    hsp6fgroup=hsp6frame.groupby('P_ID');
    psoonlydrugs=pandas.read_excel('Biologics_Humira_HS-PSO_Program_Settings.xlsx', \
                                   sheet_name='PsoriasisOnlyDrugs');
    for patient_id, patient_stats in hsp6fgroup:
        tempframe=patient_stats;
        psopositive=False;
        for molecule in tempframe['combomol']:
            for drug in psoonlydrugs['PsoriasisOnlyDrugs']:
                if str(drug) in str(molecule):
                    psopositive=True;
        for product in tempframe['prod']:
            for drug in psoonlydrugs['PsoriasisOnlyDrugs']:
                if str(drug) in str(product):
                    psopositive=True;
        if psopositive:
            print("PsO patient No."+str(patient_id)+" detected.");
            pso4frame=pandas.concat([pso4frame, tempframe]);
        else:
            print("HS probable patient No."+str(patient_id)+" detected.");
            hsp7frame=pandas.concat([hsp7frame, tempframe]);
    print("Concluding Psoriasis-only drug check on comedication.");
    
    # ---------------- PHASE III: iii. i.) 3 month Injection Basis ---------------
    
    print("Initializing Tri-monthly injection based classification.");
    hsp8_1frame=hsp8_2frame=newhspframe=newpsoframe=newhsframe=emptyframe7;
    hsp1fgroup=hsp1frame.groupby('P_ID');
    for patient_id, patient_stats in hsp1fgroup:
        plus3months=False;
        tempframe=patient_stats;
        tempframe=tempframe.sort_values(by=['TransactionDate'], ascending=True);
        i=0;
        for molecule in tempframe['combomol']:
            if 'ADALIMUMAB' in str(molecule) \
            and tempframe['ProdTransaction'].iloc[[i]].item() == 'New':
                firstdate=tempframe['TransactionDate'].iloc[[i]].item();
                transdate=firstdate;
                transpydate=datetime.date(int(transdate/10000), \
                                          int((transdate%10000)/100), \
                                          int(transdate%100));
                firstsasdate=DatePythontoSAS(transpydate);
                #break;
            i+=1;
        i=0;
        for buydate in tempframe['TransactionDate']:
            itransdate=buydate;
            itranspydate=datetime.date(int(itransdate/10000), \
                                       int((itransdate%10000)/100), \
                                       int(itransdate%100));
            isasdate=DatePythontoSAS(itranspydate);
            if 'ADALIMUMAB' in str(tempframe['combomol'].iloc[[i]].item()) \
            and int(isasdate) > int(firstsasdate + 90):
                plus3months=True;
            i+=1;
        if plus3months:
            hsp8_1frame=pandas.concat([hsp8_1frame, tempframe]);
        else:
            hsp8_2frame=pandas.concat([hsp8_2frame, tempframe]);
    
    hsp8_1fgroup=hsp8_1frame.groupby('P_ID');
    for patient_id, patient_stats in hsp8_1fgroup:
        tempframe=patient_stats;
        tempframe=tempframe.sort_values(by=['TransactionDate'], ascending=True);
        i=0;
        for molecule in tempframe['combomol']:
            if 'ADALIMUMAB' in str(molecule) \
            and tempframe['ProdTransaction'].iloc[[i]].item() == 'New':
                firstdate=tempframe['TransactionDate'].iloc[[i]].item();
                transdate=firstdate;
                transpydate=datetime.date(int(transdate/10000), \
                                          int((transdate%10000)/100), \
                                          int(transdate%100));
                firstsasdate=DatePythontoSAS(transpydate);
                #break;
            i+=1;
        sumofunits=0;
        i=0;
        for molecule in tempframe['combomol']:
            itransdate=tempframe['TransactionDate'].iloc[[i]].item();
            itranspydate=datetime.date(int(itransdate/10000), \
                                       int((itransdate%10000)/100), \
                                       int(itransdate%100));
            isasdate=DatePythontoSAS(itranspydate);
            if 'ADALIMUMAB' in str(molecule) \
            and 0 <= int(isasdate - firstsasdate) <= 90:
                sumofunits+=int(int(tempframe['PSIZE'].iloc[[i]].item())*int(tempframe['str_unit'].iloc[[i]].item())*int(tempframe['Units'].iloc[[i]].item()))/80;
            i+=1;
        if int(sumofunits) < 5:
            print("PsO patient No."+str(patient_id)+" detected.");
            newpsoframe=pandas.concat([newpsoframe, tempframe]);
        elif int(sumofunits) == 5:
            print("PsO patient No."+str(patient_id)+" detected.");
            newhspframe=pandas.concat([newhspframe, tempframe]);
        else:
            print("HS patient No."+str(patient_id)+" detected.");
            newhsframe=pandas.concat([newhsframe, tempframe]);
    print("Concluding Tri-monthly injection based classification.");
    
    #To Phase IV Therapy Duration
    
    prefinalframe=emptyframe7;
    
    hs1fgroup=hs1frame.groupby('P_ID');
    hs2fgroup=hs2frame.groupby('P_ID');
    hs3fgroup=hs3frame.groupby('P_ID');
    pso1fgroup=pso1frame.groupby('P_ID');
    pso2fgroup=pso2frame.groupby('P_ID');
    pso3fgroup=pso3frame.groupby('P_ID');
    pso4fgroup=pso4frame.groupby('P_ID');
    hsp7fgroup=hsp7frame.groupby('P_ID');
    
    for patient_id_1, patient_stats_1 in hsp1fgroup:
        tempframe=patient_stats_1;
        for patient_id_2, patient_stats_2 in hs1fgroup:
            if patient_id_1 == patient_id_2:
                tempframe['Flag_Till_Now']='HS';
                break;
        for patient_id_2, patient_stats_2 in hs2fgroup:
            if patient_id_1 == patient_id_2:
                tempframe['Flag_Till_Now']='HS';
                break;
        for patient_id_2, patient_stats_2 in hs3fgroup:
            if patient_id_1 == patient_id_2:
                tempframe['Flag_Till_Now']='HS';
                break;
        for patient_id_2, patient_stats_2 in pso1fgroup:
            if patient_id_1 == patient_id_2:
                tempframe['Flag_Till_Now']='PSO';
                break;
        for patient_id_2, patient_stats_2 in pso2fgroup:
            if patient_id_1 == patient_id_2:
                tempframe['Flag_Till_Now']='PSO';
                break;
        for patient_id_2, patient_stats_2 in pso3fgroup:
            if patient_id_1 == patient_id_2:
                tempframe['Flag_Till_Now']='PSO';
                break;
        for patient_id_2, patient_stats_2 in pso4fgroup:
            if patient_id_1 == patient_id_2:
                tempframe['Flag_Till_Now']='PSO';
                break;
        for patient_id_2, patient_stats_2 in hsp7fgroup:
            if patient_id_1 == patient_id_2:
                tempframe['Flag_Till_Now']='Unclassified';
                break;
        prefinalframe=pandas.concat([prefinalframe, tempframe]);
    
    hs1frame['Flag_Till_Now']=hs2frame['Flag_Till_Now']=hs3frame['Flag_Till_Now']='HS';
    hsp7frame['Flag_Till_Now']='Unclassified';
    pso1frame['Flag_Till_Now']=pso2frame['Flag_Till_Now']=pso3frame['Flag_Till_Now']=\
    pso4frame['Flag_Till_Now']='PSO';
    
    emptyframe8=pandas.DataFrame(columns=['Batch_ID', 'Sequence', 'BasketNumber', \
                                          'RX_ID', 'TransactionType', 'P_Gender', \
                                          'P_Loyalty', 'DCI_VOS', 'combomol', \
                                          'DDMS_PHA', 'TransactionDate', 'FCC', \
                                          'Units', 'ShortCode','P_ID', \
                                          'TransactionMonth', 'atc', 'prod', \
                                          'pack', 'str_unit','str_meas', 'gene', \
                                          'DoctorSpeciality', 'CU', 'PSIZE', \
                                          'psize', 'numdax', 'DoctorClass', \
                                          'PatientLevelClass', 'ProdTransaction', \
                                          'Flag_Till_Now', 'New_Flag']);
    
    finalframe=semifinalframe=iiiframe=excludeframe=nhsp2frame=emptyframe8;
    preffgroup=prefinalframe.groupby('P_ID');
    
    nhsfgroup=newhsframe.groupby('P_ID');
    nhspfgroup=newhspframe.groupby('P_ID');
    npsofgroup=newpsoframe.groupby('P_ID');
    hs8_2group=hsp8_2frame.groupby('P_ID');
    
    for patient_id_3, patient_stats_3 in preffgroup:
        tempframe=patient_stats_3;
        for patient_id_4, patient_stats_4 in nhsfgroup:
            if patient_id_3 == patient_id_4:
                tempframe['New_Flag']='HS';
                break;
        for patient_id_4, patient_stats_4 in nhspfgroup:
            if patient_id_3 == patient_id_4:
                #if tempframe['Flag_Till_Now'].iloc[[0]].item() == 'Unclassified':
                    #tempframe['New_Flag']='PSO';
                #else:
                    #tempframe['New_Flag']='Possible HS';
                tempframe['New_Flag']='Possible HS';
                nhsp2frame=pandas.concat([nhsp2frame, tempframe]);
                break;
        for patient_id_4, patient_stats_4 in npsofgroup:
            if patient_id_3 == patient_id_4:
                tempframe['New_Flag']='PSO';
                break;
        for patient_id_4, patient_stats_4 in hs8_2group:
            if patient_id_3 == patient_id_4:
                tempframe['New_Flag']='';
                break;
        semifinalframe=pandas.concat([semifinalframe, tempframe]);
    
    newhsframe['New_Flag']='HS';
    newhspframe['New_Flag']='Possible HS';          # nhsp2frame is an updated version. 
    newpsoframe['New_Flag']='PSO';
    hsp8_2frame['New_Flag']='';           
    
    print('Dataframes ready for NEXT phase.');
    
    #DEPRECATED ANALOGY
    
    '''oldhpframe['Next_Flag']=nonhumirafilterframe['Next_Flag']=gframe['Next_Flag']=\
    rframe['Next_Flag']=uframe['Next_Flag']=eframe['Next_Flag']='Unknown';
    semifinalframe['Next_Flag']=semifinalframe['Flag_Till_Now'];
    
    sffgroup=semifinalframe.groupby('Flag_Till_Now');
    for flag_till_now, patient_stats_9 in sffgroup:
        tempframe=patient_stats_9;
        if flag_till_now == 'HS':
            tfgroup=tempframe.groupby('New_Flag');
            for new_flag, patient_stats_10 in tfgroup:
                tempframe2 =patient_stats_10;
                if new_flag == 'PSO':
                    tempframe2['Next_Flag']='PSO';
                semifinalframe2=pandas.concat([semifinalframe2, tempframe2]);
        else:
            semifinalframe2=pandas.concat([semifinalframe2, tempframe]);'''
    
    # SUMMARY FILE CREATION
    
    summaryframe=pandas.DataFrame(columns=['P_ID', 'ANTI-TNF', 'PACKS_1ST_MONTH', \
                                           'ANTIBIOTIC', 'METH_CICLO', \
                                           'ISO-BETADINE', 'COMEDICATION', \
                                           'OBSERVATION_PERIOD', 'PANEL_PERIOD', 'MAX HUMIRA GAP', \
                                           'PACKS_12_MONTHS', 'TOTAL_PACKS', 'TOTAL_MG_CONSUMED', 'TOTAL TREATMENT IN 40MG PACKS', \
                                           '1ST_TRANS_DATE_ON_HUMIRA', \
                                           'LAST_TRANS_DATE_ON_HUMIRA', \
                                           'PATPROFIL', \
                                           'SINGLE_PURCHASE', 'SHARE OF PACKS PER WEEK', 'SHARE OF TREATMENT PER WEEK', 'PROPOSED_FLAG']);
    
    print("Importing Patient Profile SAS dataset...");
    #importedsasfile3=input("Enter the filename of Patient Profile SAS dataset to be imported(without ext.):")
    importedsasfile3=filelocs['Address'][2];
    patprofilframe=pandas.read_sas(importedsasfile3, format='sas7bdat', encoding='iso-8859-1');
    
    hsplist=[hs1frame, hs2frame, hs3frame, hsp7frame, pso1frame, \
             pso2frame, pso3frame, pso4frame];
    
    i=k=0;
    for patientframe in hsplist:
        pfgroup=patientframe.groupby('P_ID');
            # P_ID
        for patient_id, patient_stats in pfgroup:
            # ANTI-TNF, ANTIBIOTIC, METH_CICLO, ISO-BETADINE or COMEDICATION
            antitnf=antibiotics=meth_ciclo=isobetadine=comedication=False
            tempframe=patient_stats;
            tempframe=tempframe.sort_values(by=['TransactionDate'], ascending=True);
            if i == 4:
                antitnf=True;
            if i == 1:
                antibiotics=True;
            if i == 6:
                meth_ciclo=True;
            if i == 2:
                isobetadine=True;
            if i == 7:
                comedication=True;
            # 1ST_TRANS_DATE_ON_HUMIRA (latest)
            j=0;
            for molecule in tempframe['combomol']:
                if 'ADALIMUMAB' in str(molecule) \
                and tempframe['ProdTransaction'].iloc[[j]].item() == 'New':
                    firstdate=tempframe['TransactionDate'].iloc[[j]].item();
                    transdate=firstdate;
                    transpydate=datetime.date(int(transdate/10000), \
                                              int((transdate%10000)/100), \
                                              int(transdate%100));
                    firstsasdate=DatePythontoSAS(transpydate);
                j+=1;
            # TOTAL_PACKS, PACKS_1ST_MONTH, PACKS_12_MONTHS, OBSERVATION_PERIOD,
            # LAST_TRANS_DATE_ON_HUMIRA and FLAG_TILL_NOW
            total_packs=0;
            first_month_packs=0;
            third_month_packs=0;
            j=0;
            for jbuydate in tempframe['TransactionDate']:
                jtransdate=jbuydate;
                jtranspydate=datetime.date(int(jtransdate/10000), \
                                           int((jtransdate%10000)/100), \
                                           int(jtransdate%100));
                jsasdate=DatePythontoSAS(jtranspydate);
                if 'ADALIMUMAB' in str(tempframe['combomol'].iloc[[j]].item()) \
                and 0 <= int(jsasdate - firstsasdate) <= 30:
                    first_month_packs+=int(int(tempframe['PSIZE'].iloc[[j]].item())*int(tempframe['str_unit'].iloc[[j]].item())*int(tempframe['Units'].iloc[[j]].item()))/80;
                if 'ADALIMUMAB' in str(tempframe['combomol'].iloc[[j]].item()) \
                and 0 <= int(jsasdate - firstsasdate) <= 84:
                    third_month_packs+=int(tempframe['Units'].iloc[[j]].item());
                if 'ADALIMUMAB' in str(tempframe['combomol'].iloc[[j]].item()) \
                and int(jsasdate) >= int(firstsasdate):
                    total_packs+=int(tempframe['Units'].iloc[[j]].item());
                if 'ADALIMUMAB' in str(tempframe['combomol'].iloc[[j]].item()) \
                or 'HUMIRA' in str(tempframe['prod'].iloc[[j]].item()):
                    lastdate=jbuydate;
                    proposed_flag=tempframe['Flag_Till_Now'].iloc[[j]].item();
                j+=1;
            # TOTAL_MG_CONSUMPTION & TOTAL TREATMENT IN 40MG PACKS
            total_mg_consumed=0;
            for index, rx in tempframe.iterrows():
                if 'ADALIMUMAB' in str(rx['combomol']):
                    rx_amount=rx['str_unit']*rx['PSIZE']*rx['Units'];
                    total_mg_consumed+=rx_amount;
            total_mg_consumed=int(total_mg_consumed);
            totaltreatmentin40mgpacks=total_mg_consumed/80;
            # PATPROFIL (dlast)
            j=0;
            dlast=0;
            for patient_id_1 in patprofilframe['P_ID']:
                if patient_id == patient_id_1:
                   dlast=patprofilframe['dlast'].iloc[[j]].item();
                j+=1;
            # OBSERVATION PERIOD
            lastpydate=datetime.date(int(lastdate/10000), \
                                     int((lastdate%10000)/100), \
                                     int(lastdate%100));
            lastsasdate=DatePythontoSAS(lastpydate);
            observation_period=round((lastsasdate-firstsasdate)/7);
            # PANEL PERIOD
            lastpydate=datetime.date(int(dlast/10000), \
                                     int((dlast%10000)/100), \
                                     int(dlast%100));
            lastsasdate=DatePythontoSAS(lastpydate);
            panel_period=round((lastsasdate-firstsasdate)/7);
            # SHARE OF TREATMENT PER WEEK
            if observation_period != 0:
                shareoftreatmentperweek=round((totaltreatmentin40mgpacks/observation_period)*100);
            else:
                shareoftreatmentperweek=-1;
            # SHARE OF PACKS PER WEEK
            if observation_period != 0:
                shareofpacksperweek=round((total_packs/observation_period)*100);
            else:
                shareofpacksperweek=-1;
            # MAXIMUM GAP BETWEEN HUMIRA RXS
            deltahumira=-1;
            j=0;
            for jbuydate in tempframe['TransactionDate']:
                if 'ADALIMUMAB' in str(tempframe['combomol'].iloc[[j]].item()) \
                or 'HUMIRA' in str(tempframe['prod'].iloc[[j]].item()):
                    if deltahumira == -1:
                        prevhumiradate=tempframe['numdax'].iloc[[j]].item();
                        deltahumira+=1;
                    else:
                        currhumiradate=tempframe['numdax'].iloc[[j]].item();
                        diffhumiradate=currhumiradate-prevhumiradate;
                        prevhumiradate=currhumiradate;
                        if diffhumiradate > deltahumira:
                            deltahumira = diffhumiradate;
                j+=1;
            deltahumira=round(deltahumira/7);
            # SINGLE PURCHASE
            single_purchase=False;
            j=0;
            for jbuydate in tempframe['TransactionDate']:
                if 'ADALIMUMAB' in str(tempframe['combomol'].iloc[[j]].item()) \
                or 'HUMIRA' in str(tempframe['prod'].iloc[[j]].item()):
                    firstpurchasedate=jbuydate;
                    break;
                j+=1;
            j=0;
            for jbuydate in tempframe['TransactionDate']:
                if 'ADALIMUMAB' in str(tempframe['combomol'].iloc[[j]].item()) \
                or 'HUMIRA' in str(tempframe['prod'].iloc[[j]].item()):
                    lastpurchasedate=jbuydate;
                j+=1;
            if firstpurchasedate == lastpurchasedate:
                single_purchase=True;
            # DATA CREATION
            patientsummary=numpy.array([int(patient_id), antitnf, first_month_packs, \
                                        antibiotics, meth_ciclo, isobetadine, \
                                        comedication, observation_period, panel_period, deltahumira, \
                                        third_month_packs, total_packs, total_mg_consumed, totaltreatmentin40mgpacks, \
                                        int(firstdate), int(lastdate), int(dlast), single_purchase, \
                                        shareofpacksperweek, shareoftreatmentperweek, proposed_flag]);
            summaryframe.loc[k]=patientsummary;
            k+=1;
        i+=1;
    print("...Created the pre-finalised summary file.");
    finalsummaryframe=pandas.DataFrame(columns=['P_ID', 'ANTI-TNF', 'PACKS_1ST_MONTH', \
                                                'ANTIBIOTIC', 'METH_CICLO', \
                                                'ISO-BETADINE', 'COMEDICATION', \
                                                'OBSERVATION_PERIOD', 'PANEL_PERIOD', 'MAX HUMIRA GAP', \
                                                'PACKS_12_MONTHS', 'TOTAL_PACKS', 'TOTAL_MG_CONSUMED', 'TOTAL TREATMENT IN 40MG PACKS', \
                                                '1ST_TRANS_DATE_ON_HUMIRA', 'LAST_TRANS_DATE_ON_HUMIRA', 'PATPROFIL', \
                                                'SINGLE_PURCHASE', 'SHARE OF PACKS PER WEEK', 'SHARE OF TREATMENT PER WEEK', 'PROPOSED_FLAG', 'UNITS_BASED_FLAG']);
    sfgroup=summaryframe.groupby('P_ID');
    sffgroup=semifinalframe.groupby('P_ID');
    for patient_id_5, patient_stats_5 in sfgroup:
        tempframe=patient_stats_5
        for patient_id_6, patient_stats_6 in sffgroup:
            tempframe2=patient_stats_6;
            if int(patient_id_5) == int(patient_id_6):
                tempframe['UNITS_BASED_FLAG']=tempframe2['New_Flag'].iloc[[0]].item();
                finalsummaryframe=pandas.concat([finalsummaryframe, tempframe]);
    finalsummaryframe2=finalsummaryframe[0:0];
    for index, rx in finalsummaryframe.iterrows():
        if rx['ANTI-TNF'] == 'True':
            rx['FINAL_FLAG']='PSO';
        if rx['ANTIBIOTIC'] == 'True':
            if int(rx['SHARE OF TREATMENT PER WEEK']) >= 45:
                rx['FINAL_FLAG']='HS';
            elif 40 <= int(rx['SHARE OF TREATMENT PER WEEK']) < 45:
                rx['FINAL_FLAG']='REVIEW';
            else:
                rx['FINAL_FLAG']='PSO';
        if rx['METH_CICLO'] == 'True':
            rx['FINAL_FLAG']='PSO';
        if rx['ISO-BETADINE'] == 'True':
            if int(rx['SHARE OF TREATMENT PER WEEK']) >= 45:
                rx['FINAL_FLAG']='HS';
            elif 40 <= int(rx['SHARE OF TREATMENT PER WEEK']) < 45:
                rx['FINAL_FLAG']='REVIEW';
            else:
                rx['FINAL_FLAG']='PSO';
        if rx['COMEDICATION'] == 'True':
            rx['FINAL_FLAG']='PSO';
        if rx['PROPOSED_FLAG'] == 'HS':
            if rx['UNITS_BASED_FLAG'] == 'HS':
                if int(rx['SHARE OF TREATMENT PER WEEK']) >= 45:
                    rx['FINAL_FLAG']='HS';
                elif 40 <= int(rx['SHARE OF TREATMENT PER WEEK']) < 45:
                    rx['FINAL_FLAG']='REVIEW';
                else:
                    rx['FINAL_FLAG']='PSO';
            elif rx['UNITS_BASED_FLAG'] == 'Possible HS':
                if int(rx['SHARE OF TREATMENT PER WEEK']) >= 45:
                    rx['FINAL_FLAG']='HS';
                elif 40 <= int(rx['SHARE OF TREATMENT PER WEEK']) < 45:
                    rx['FINAL_FLAG']='REVIEW';
                else:
                    rx['FINAL_FLAG']='PSO';
            elif rx['UNITS_BASED_FLAG'] == 'PSO':
                if int(rx['SHARE OF TREATMENT PER WEEK']) >= 45:
                    rx['FINAL_FLAG']='HS';
                elif 40 <= int(rx['SHARE OF TREATMENT PER WEEK']) < 45:
                    rx['FINAL_FLAG']='REVIEW';
                else:
                    rx['FINAL_FLAG']='PSO';
            else:
                if int(rx['SHARE OF TREATMENT PER WEEK']) >= 45:
                    rx['FINAL_FLAG']='HS';
                elif 40 <= int(rx['SHARE OF TREATMENT PER WEEK']) < 45:
                    rx['FINAL_FLAG']='REVIEW';
                else:
                    rx['FINAL_FLAG']='PSO';
        if rx['PROPOSED_FLAG'] == 'PSO':
            rx['FINAL_FLAG']='PSO';
        if rx['PROPOSED_FLAG'] == 'Unclassified':
            if rx['UNITS_BASED_FLAG'] == 'PSO':
                if int(rx['SHARE OF TREATMENT PER WEEK']) >= 45:
                    rx['FINAL_FLAG']='HS';
                elif 40 <= int(rx['SHARE OF TREATMENT PER WEEK']) < 45:
                    rx['FINAL_FLAG']='REVIEW';
                else:
                    rx['FINAL_FLAG']='PSO';
            else:
                rx['FINAL_FLAG']='PSO';
        print("Assigning the Final flag to the patients...");
        finalsummaryframe2=finalsummaryframe2.append(rx);
    #semifinalframe['Final_Flag']='';
    
    '''
        for patientframe in hsplist2:
            pfgroup=patientframe.groupby('P_ID');
            for patient_id_6, patient_stats_6 in pfgroup:
                if patient_id_5 == patient_id_6:
                    tempframe=patient_stats_5;
                    if i == 0:
                        tempframe['NEW_FLAG']='HS';
                        finalsummaryframe=pandas.concat([finalsummaryframe, tempframe]);
                        break;
                    if i == 1:
                        finalsummaryframe=pandas.concat([finalsummaryframe, tempframe]);
                        tempframe['NEW_FLAG']='Possible HS';
                        break;
                    if i == 2:
                        finalsummaryframe=pandas.concat([finalsummaryframe, tempframe]);
                        tempframe['NEW_FLAG']='PSO';
                        break;
                    if i == 3:
                        finalsummaryframe=pandas.concat([finalsummaryframe, tempframe]);
                        tempframe['NEW_FLAG']='';
                        break;
            i+=1;
    '''
    print("Exporting the Summary file to excel spreadsheets...");
    finalsummaryframe2[['P_ID', 'PACKS_1ST_MONTH', 'OBSERVATION_PERIOD', 'PANEL_PERIOD', \
                        'MAX HUMIRA GAP', 'PACKS_12_MONTHS', 'TOTAL_PACKS', 'TOTAL_MG_CONSUMED', \
                        'TOTAL TREATMENT IN 40MG PACKS', '1ST_TRANS_DATE_ON_HUMIRA', \
                        'LAST_TRANS_DATE_ON_HUMIRA', 'PATPROFIL', 'SHARE OF PACKS PER WEEK', \
                        'SHARE OF TREATMENT PER WEEK']]\
                        =finalsummaryframe2[['P_ID', 'PACKS_1ST_MONTH', 'OBSERVATION_PERIOD', 'PANEL_PERIOD', \
                                             'MAX HUMIRA GAP', 'PACKS_12_MONTHS', 'TOTAL_PACKS', 'TOTAL_MG_CONSUMED', \
                                             'TOTAL TREATMENT IN 40MG PACKS', '1ST_TRANS_DATE_ON_HUMIRA', \
                                             'LAST_TRANS_DATE_ON_HUMIRA', 'PATPROFIL', 'SHARE OF PACKS PER WEEK', \
                                             'SHARE OF TREATMENT PER WEEK']].apply(pandas.to_numeric, errors='coerce');
    finalsummaryframe2=finalsummaryframe2.sort_values(by=['P_ID'], ascending=True);
    finalsummaryframe2.to_excel("Final_Summary_File.xlsx", sheet_name="Summary_File", index=False);
    
    # HUMIRA SPLIT FILE CREATION
    
    humira_mg_split=pandas.DataFrame(columns=['P_ID', 'prod', 'numdax', \
                                              'TransactionDate', 'StartDateofHumira', \
                                              'LastDateofHumira', 'str_unit', \
                                              'PSIZE', 'Units']);
    
    hsp1fgroup=hsp1frame.groupby('P_ID');
    for patient_id_9, patient_stats_9 in sffgroup:
        tempframe=patient_stats_9;
        tempframe=tempframe.sort_values(by=['numdax'], ascending=False);
        tempframe['LastDateofHumira']=tempframe['TransactionDate'].iloc[[0]].item();
        tempframe=tempframe.sort_values(by=['numdax'], ascending=True);
        tempframe['StartDateofHumira']=tempframe['TransactionDate'].iloc[[0]].item();
        tempframe=tempframe[['P_ID', 'prod', 'numdax', 'TransactionDate', \
                             'StartDateofHumira', 'LastDateofHumira', 'str_unit', \
                             'PSIZE', 'Units']];
        humira_mg_split=pandas.concat([humira_mg_split, tempframe]);
    
    humira_mg_split.to_excel("HUMIRA_MG_SPLIT.xlsx", sheet_name="Humira_mg_Split", index=False);
    
    # REVIEW OF THE SUMMARY FILE:
    
    input("***** Review the Summary File and then press ENTER to continue *****");
    finalsummaryframe3=pandas.read_excel("Final_Summary_File.xlsx", sheet_name='Summary_File');
    finalsummaryframe3[['UNITS_BASED_FLAG', 'FINAL_FLAG']]=finalsummaryframe3[['UNITS_BASED_FLAG','FINAL_FLAG']].fillna(value='');
    semifinalframe['Final_Flag']='';
    i=0;
    for indx, rx in semifinalframe.iterrows():
        for indx2, patient_entry in finalsummaryframe3.iterrows():
            if rx['P_ID'] == int(patient_entry['P_ID']):
                semifinalframe['Final_Flag'].iloc[[i]]=patient_entry['FINAL_FLAG'];
                print("Patient logged in the main dataframe.");
        i+=1;
    
    # EXPORTING DATASETS FOR QC:
    
    if not os.path.isdir('QC Datasets'):
        os.makedirs('QC Datasets');
    dir_path = os.path.dirname(os.path.realpath(__file__));
    gframe.to_excel(dir_path+"/QC Datasets/"+"Gastro Patients(gframe).xlsx", sheet_name="g", index=False);
    rframe.to_excel(dir_path+"/QC Datasets/"+"Rheumato Patients(rframe).xlsx", sheet_name="hsp", index=False);
    dframe.to_excel(dir_path+"/QC Datasets/"+"Dermato Patients(dframe).xlsx", sheet_name="hsp", index=False);
    eframe.to_excel(dir_path+"/QC Datasets/"+"Irrelevant Patients(eframe).xlsx", sheet_name="hsp", index=False);
    hsp1frame.to_excel(dir_path+"/QC Datasets/"+"Indication Split HSP Starting Set 1 (hsp1frame).xlsx", sheet_name="hsp", index=False);
    hsp2frame.to_excel(dir_path+"/QC Datasets/"+"Proposed Flag Split HSP Remaining Set 2 (hsp2frame).xlsx", sheet_name="hsp", index=False);
    hsp3frame.to_excel(dir_path+"/QC Datasets/"+"Proposed Flag Split HSP Remaining Set 3 (hsp3frame).xlsx", sheet_name="hsp", index=False);
    hsp4frame.to_excel(dir_path+"/QC Datasets/"+"Proposed Flag Split HSP Remaining Set 4 (hsp4frame).xlsx", sheet_name="hsp", index=False);
    hsp5frame.to_excel(dir_path+"/QC Datasets/"+"Proposed Flag Split HSP Remaining Set 5 (hsp5frame).xlsx", sheet_name="hsp", index=False);
    hsp6frame.to_excel(dir_path+"/QC Datasets/"+"Proposed Flag Split HSP Remaining Set 6 (hsp6frame).xlsx", sheet_name="hsp", index=False);
    hsp7frame.to_excel(dir_path+"/QC Datasets/"+"Proposed Flag Split HSP Unclassified Set 7 (hsp7frame).xlsx", sheet_name="hsp", index=False);
    hs1frame.to_excel(dir_path+"/QC Datasets/"+"Proposed Flag Split HS 30 day Inj Set 3 (hs1frame).xlsx", sheet_name="hs", index=False);
    hs2frame.to_excel(dir_path+"/QC Datasets/"+"Proposed Flag Split HS Antibiotic Set 4 (hs2frame).xlsx", sheet_name="hs", index=False);
    hs3frame.to_excel(dir_path+"/QC Datasets/"+"Proposed Flag Split HS Isobetadine Set 6 (hs3frame).xlsx", sheet_name="hs", index=False);
    pso1frame.to_excel(dir_path+"/QC Datasets/"+"Proposed Flag Split PSO Anti-TNF Set 2 (pso1frame).xlsx", sheet_name="pso", index=False);
    pso2frame.to_excel(dir_path+"/QC Datasets/"+"Proposed Flag Split PSO 30 days Inj Set 2 (pso2frame).xlsx", sheet_name="pso", index=False);
    pso3frame.to_excel(dir_path+"/QC Datasets/"+"Proposed Flag Split PSO Meth-Ciclo Set 5 (pso3frame).xlsx", sheet_name="pso", index=False);
    pso4frame.to_excel(dir_path+"/QC Datasets/"+"Proposed Flag Split PSO Comedication Set 7 (pso4frame).xlsx", sheet_name="pso", index=False);
    hsp8_1frame.to_excel(dir_path+"/QC Datasets/"+"Unit-based Flag Split HSP 84 day+ obser Set 2 (hsp8_1frame).xlsx", sheet_name="hsp", index=False);
    hsp8_2frame.to_excel(dir_path+"/QC Datasets/"+"Unit-based Flag Split Blank 84 day- obser Set 2 (hsp8_2frame).xlsx", sheet_name="hsp", index=False);
    newhspframe.to_excel(dir_path+"/QC Datasets/"+"Unit-based Flag Split HSP 84 day Inj Set 3 (newhspframe).xlsx", sheet_name="hsp", index=False);
    newhsframe.to_excel(dir_path+"/QC Datasets/"+"Unit-based Flag Split HS 84 day Inj Set 3 (newhsframe).xlsx", sheet_name="hs", index=False);
    newpsoframe.to_excel(dir_path+"/QC Datasets/"+"Unit-based Flag Split PSO 84 day Inj Set 3 (newpsoframe).xlsx", sheet_name="pso", index=False);
    
    # QC CHECK SUMMARY FILE
    
    qcchecksummary=pandas.DataFrame(columns=['Dataset', 'Number_of_Patients', 'Number of rx per units'], \
                                    data=[['Biologics dataset', NumofPatients(sasframe), len(sasframe)], \
                                          ['Gastro Patients', NumofPatients(gframe), len(gframe)], \
                                          ['Rheumato Patients', NumofPatients(rframe), len(rframe)], \
                                          ['Dermato Patients', NumofPatients(dframe), len(dframe)], \
                                          ['Irrelevant Patients', NumofPatients(eframe), len(eframe)], \
                                          ['Total of subsets', NumofPatients(sasframe), len(sasframe)], \
                                          ['Indication Split HSP Starting 1', NumofPatients(hsp1frame), len(hsp1frame)], \
                                          ['Proposed Flag HSP Remaining 2', NumofPatients(hsp2frame), len(hsp2frame)], \
                                          ['Proposed Flag PSO Anti-TNF 2', NumofPatients(pso1frame), len(pso1frame)], \
                                          ['Proposed Flag HSP Remaining 3', NumofPatients(hsp3frame), len(hsp3frame)], \
                                          ['Proposed Flag HSP Master 3.5', NumofPatients(hsp3masterframe), len(hsp3masterframe)], \
                                          ['Proposed Flag HS 30 day Inj 3', NumofPatients(hs1frame), len(hs1frame)], \
                                          ['Proposed Flag PSO 30 day Inj 3', NumofPatients(pso2frame), len(pso2frame)], \
                                          ['Proposed Flag HSP Remaining 4', NumofPatients(hsp4frame), len(hsp4frame)], \
                                          ['Proposed Flag HS Antibiotic 4', NumofPatients(hs2frame), len(hs2frame)], \
                                          ['Proposed Flag HSP Remaining 5', NumofPatients(hsp5frame), len(hsp5frame)], \
                                          ['Proposed Flag PSO Meth-Ciclo 3', NumofPatients(pso3frame), len(pso3frame)], \
                                          ['Proposed Flag HSP Remaining 6', NumofPatients(hsp6frame), len(hsp6frame)], \
                                          ['Proposed Flag HS Isobetadine 6', NumofPatients(hs3frame), len(hs3frame)], \
                                          ['Proposed Flag HSP Unclassified 7', NumofPatients(hsp7frame), len(hsp7frame)], \
                                          ['Proposed Flag PSO Comedication 7', NumofPatients(pso4frame), len(pso4frame)], \
                                          ['Unit-based Flag HSP 84 day+ obser 2', NumofPatients(hsp8_1frame), len(hsp8_1frame)], \
                                          ['Unit-based Flag Blank 84 day- obser 2', NumofPatients(hsp8_2frame), len(hsp8_2frame)], \
                                          ['Unit-based Flag HSP 84 day Inj 3', NumofPatients(newhspframe), len(newhspframe)], \
                                          ['Unit-based Flag HS 84 day Inj 3', NumofPatients(newhsframe), len(newhsframe)], \
                                          ['Unit-based Flag PSO 84 day Inj 3', NumofPatients(newpsoframe), len(newpsoframe)]]);
    
    qcchecksummary.to_excel(dir_path+"/QC Datasets/"+"QC_Check_Summary.xlsx", sheet_name="QC", index=False);
    
    # FINAL FRAME CREATION
    
    oldhpframe['Final_Flag']='PSO';     # Old Patients 'New' on Humira are auto 'PSO'.
    nonhumirafilterframe['Final_Flag']=gframe['Final_Flag']=\
    rframe['Final_Flag']=uframe['Final_Flag']=eframe['Final_Flag']='Unknown';
    finalframe=pandas.concat([semifinalframe, oldhpframe, nonhumirafilterframe, gframe, rframe, uframe, eframe]);
    '''
    print('Excluding unclassified patients.');
    ffgroup=finalframe.groupby('Next_Flag');
    for next_flag, patient_stats in ffgroup:
        tempframe=patient_stats
        if next_flag != 'Unclassified':
            iiiframe=pandas.concat([iiiframe, tempframe]);
        else:
            excludeframe=pandas.concat([excludeframe, tempframe]);
    '''
    iiiframe=finalframe;
    
    # ------------------------ PHASE IV: Pack Numbering --------------------------
    
    emptyframe4=pandas.DataFrame(columns=['Batch_ID', 'Sequence', 'BasketNumber', \
                                          'RX_ID', 'TransactionType', 'P_Gender', \
                                          'P_Loyalty', 'DCI_VOS', 'combomol', \
                                          'DDMS_PHA', 'TransactionDate', 'FCC', \
                                          'Units', 'ShortCode','P_ID', \
                                          'TransactionMonth', 'atc', 'prod', \
                                          'pack', 'str_unit','str_meas', 'gene', \
                                          'DoctorSpeciality', 'CU', 'PSIZE', \
                                          'psize', 'numdax', 'DoctorClass', \
                                          'PatientLevelClass', 'ProdTransaction', \
                                          'Flag_Till_Now', 'New_Flag', 'Final_Flag', 'PackNumber']);
    ivframe=emptyframe4;
    #iiiframe=pandas.read_excel("IIIFrame.xlsx");
    iiiframe=iiiframe.sort_values(by=['P_ID', 'numdax', 'combomol'], ascending=[True, True, True]);
    print("Dataframe sorted by Patient ID, Transaction date and Molecule.");
    iiiframe['PackNumber']=0;
    iiiflength=len(iiiframe);
    if cpu_count() < 2:
        partitions=1;
    elif cpu_count() < 4:
        partitions=2;
    elif cpu_count() < 8:
        partitions=4;
    else:
        partitions=8;
    processedpartitionframelist=[];
    partitionframelist=PartitionDataFramebyCores(iiiframe, partitions);
    pool=Pool();
    processedpartitionframelist.append(pool.map(PackNumbering, partitionframelist));
    pool.close();
    pool.join();
    for i in range(partitions):
        ivframe=pandas.concat([ivframe, processedpartitionframelist[0][i]]);
    ivframe.to_excel("IVFrame(III).xlsx", sheet_name="IVFrame", index=False);
    
    print("Program execution ended at"+str(datetime.datetime.now().time()));
    
    # ----------------------- PHASE IV: Therapy Duration -------------------------
    
    therapyduration=pandas.read_excel('Biologics_Humira_HS-PSO_Program_Settings.xlsx', sheet_name='TreatmentDuration');
    emptyframe5=pandas.DataFrame(columns=['Batch_ID', 'Sequence', 'BasketNumber', \
                                          'RX_ID', 'TransactionType', 'P_Gender', \
                                          'P_Loyalty', 'DCI_VOS', 'combomol', \
                                          'DDMS_PHA', 'TransactionDate', 'FCC', \
                                          'Units', 'ShortCode','P_ID', \
                                          'TransactionMonth', 'atc', 'prod', \
                                          'pack', 'str_unit','str_meas', 'gene', \
                                          'DoctorSpeciality', 'CU', 'PSIZE', \
                                          'psize', 'numdax', 'DoctorClass', \
                                          'PatientLevelClass', 'ProdTransaction', \
                                          'Flag_Till_Now', 'New_Flag',\
                                          'Final_Flag', 'PackNumber', 'RXDuration']);
    ivframe=ivframe.reset_index(drop=True);
    ivframe2=emptyframe5;
    ivfgroup=ivframe.groupby('combomol');
    i=0;
    for index_2, row_2 in therapyduration.iterrows():
        tempframe=row_2
        rxd=1;
        for molecule, patient_stats_td_1 in ivfgroup:
            if molecule == tempframe['MOLECULE']:
                print(str(molecule)+" group found.")
                iterframe=patient_stats_td_1;
                ifgroup=iterframe.groupby('Final_Flag');
                for indication, patient_stats_td_2 in ifgroup:
                    if indication == tempframe['INDICATION_STAT'] or tempframe.isnull()['INDICATION_STAT']:
                        print(str(indication)+" group found.")
                        iterframe2=patient_stats_td_2;
                        if2group=iterframe2.groupby('PatientLevelClass');
                        for patientlevelclass, patient_stats_td_3 in if2group:
                            if patientlevelclass == tempframe['DOCTOR_CLASS'] or tempframe.isnull()['DOCTOR_CLASS']:
                                print(str(patientlevelclass)+" group found.");
                                iterframe3=patient_stats_td_3;
                                if3group=iterframe3.groupby('str_unit');
                                for str_unit, patient_stats_td_4 in if3group:
                                    if float(str_unit) == float(tempframe['STR_UNIT_RESTRICTION']) or tempframe.isnull()['STR_UNIT_RESTRICTION']:
                                        print("Strength units of "+str(str_unit)+" found.");
                                        iterframe4=patient_stats_td_4;
                                        if4group=iterframe4.groupby('PSIZE');
                                        for PSIZE, patient_stats_td_5 in if4group:
                                            if PSIZE == float(tempframe['PSIZE_RESTRICTION']) or tempframe.isnull()['PSIZE_RESTRICTION']:
                                                print("Packet size of "+str(PSIZE)+" found.");
                                                iterframe5=patient_stats_td_5;
                                                iterframe5['RXDuration']=1;
                                                induction1=False;
                                                induction2=False;
                                                maintenance=False;
                                                induction1frame=induction2frame=maintenanceframe=emptyframe5;
                                                if not tempframe.isnull()['INDUCTION_PHASE_II']:
                                                    induction2=True;
                                                    induction1=True;
                                                    maintenance=True;
                                                    induction2packnum=int(tempframe['INDUCTION_PHASE_II']);
                                                    induction1packnum=int(tempframe['INDUCTION_PHASE']);
                                                elif not tempframe.isnull()['INDUCTION_PHASE']:
                                                    induction1=True;
                                                    maintenance=True;
                                                    induction1packnum=int(tempframe['INDUCTION_PHASE']);
                                                else:
                                                    maintenance=True;
                                                if induction2:
                                                    if5group=iterframe5.groupby('PackNumber');
                                                    for packnum, patient_stats_td_6 in if5group:
                                                        if packnum <= induction1packnum:
                                                            iterframe6=patient_stats_td_6;
                                                            induction1frame=pandas.concat([induction1frame, iterframe6]);
                                                        elif induction1packnum < packnum <= induction2packnum:
                                                            iterframe6=patient_stats_td_6;
                                                            induction2frame=pandas.concat([induction2frame, iterframe6]);
                                                        else:
                                                            iterframe6=patient_stats_td_6;
                                                            maintenanceframe=pandas.concat([maintenanceframe, iterframe6]);
                                                elif induction1:
                                                    if5group=iterframe5.groupby('PackNumber');
                                                    for packnum, patient_stats_td_6_2 in if5group:
                                                        if packnum <= induction1packnum:
                                                            iterframe6=patient_stats_td_6_2;
                                                            induction1frame=pandas.concat([induction1frame, iterframe6]);
                                                        else:
                                                            iterframe6=patient_stats_td_6_2;
                                                            maintenanceframe=pandas.concat([maintenanceframe, iterframe6]);
                                                elif maintenance:
                                                    maintenanceframe=pandas.concat([maintenanceframe, iterframe5]);
                                                if induction1 and len(induction1frame) != 0:
                                                    strformula=tempframe['INDUCTION_FORMULA'];
                                                    if '*' in str(strformula):
                                                        varformula=map(str, strformula.split('*'));
                                                        for var in varformula:
                                                            if var=='PSIZE':
                                                                induction1frame['RXDuration']=induction1frame.apply(lambda row: row['RXDuration']*row['PSIZE'], axis=1);
                                                            elif var=='str_unit':
                                                                induction1frame['RXDuration']=induction1frame.apply(lambda row: row['RXDuration']*row['str_unit'], axis=1);
                                                            elif var=='Units':
                                                                induction1frame['RXDuration']=induction1frame.apply(lambda row: row['RXDuration']*row['Units'], axis=1);
                                                            else:
                                                                induction1frame['RXDuration']=induction1frame.apply(lambda row: row['RXDuration']*float(var), axis=1);
                                                    else:
                                                        induction1frame['RXDuration']=float(strformula);
                                                    ivframe2=pandas.concat([ivframe2, induction1frame]);
                                                    print("RX Duration for Induction calculated.")
                                                if induction2 and len(induction2frame) != 0:
                                                    strformula=tempframe['INDUCTION_FORMULA_II'];
                                                    if '*' in str(strformula):
                                                        varformula=map(str, strformula.split('*'));
                                                        for var in varformula:
                                                            if var=='PSIZE':
                                                                induction2frame['RXDuration']=induction2frame.apply(lambda row: row['RXDuration']*row['PSIZE'], axis=1);
                                                            elif var=='str_unit':
                                                                induction2frame['RXDuration']=induction2frame.apply(lambda row: row['RXDuration']*row['str_unit'], axis=1);
                                                            elif var=='Units':
                                                                induction2frame['RXDuration']=induction2frame.apply(lambda row: row['RXDuration']*row['Units'], axis=1);
                                                            else:
                                                                induction2frame['RXDuration']=induction2frame.apply(lambda row: row['RXDuration']*float(var), axis=1);
                                                    else:
                                                        induction2frame['RXDuration']=float(strformula);
                                                    ivframe2=pandas.concat([ivframe2, induction2frame]);
                                                    print("RX Duration for Induction II calculated.")
                                                if maintenance and len(maintenanceframe) != 0:
                                                    strformula=tempframe['MAINTENANCE_FORMULA'];
                                                    if '*' in str(strformula):
                                                        varformula=map(str, strformula.split('*'));
                                                        for var in varformula:
                                                            if var=='PSIZE':
                                                                maintenanceframe['RXDuration']=maintenanceframe.apply(lambda row: row['RXDuration']*row['PSIZE'], axis=1);
                                                            elif var=='str_unit':
                                                                maintenanceframe['RXDuration']=maintenanceframe.apply(lambda row: row['RXDuration']*row['str_unit'], axis=1);
                                                            elif var=='Units':
                                                                maintenanceframe['RXDuration']=maintenanceframe.apply(lambda row: row['RXDuration']*row['Units'], axis=1);
                                                            else:
                                                                maintenanceframe['RXDuration']=maintenanceframe.apply(lambda row: row['RXDuration']*float(var), axis=1);
                                                    else:
                                                        maintenanceframe['RXDuration']=float(strformula);
                                                    ivframe2=pandas.concat([ivframe2, maintenanceframe]);
                                                    print("RX Duration for Maintenance calculated.");
                                            else:
                                                print("Packet size mismatch.");
                                    else:
                                        print("Strength unit mismatch.");
                            else:
                                print("Doctor class not found.");
                    else:
                        print("Indication group not found.")
                        continue;
            else:
                continue;
        i+=1;
        print('THERAPY DURATION CRITERION NO.'+str(i)+' COMPLETED.');
    
    columnsinorder=['combomol','prod',  'P_ID', 'TransactionDate', 'DoctorClass', \
                    'PatientLevelClass', 'ProdTransaction', 'Flag_Till_Now', \
                    'New_Flag', 'Final_Flag', 'PackNumber', 'RXDuration', 'Batch_ID', \
                    'Sequence', 'BasketNumber', \
                    'RX_ID', 'TransactionType', 'P_Gender', \
                    'P_Loyalty', 'DCI_VOS', \
                    'DDMS_PHA', 'FCC', \
                    'Units', 'ShortCode',\
                    'TransactionMonth', 'atc', \
                    'pack', 'str_unit','str_meas', 'gene', \
                    'DoctorSpeciality', 'CU', 'PSIZE', \
                    'psize', 'numdax'];
    ivframe2=ivframe2[columnsinorder];
    ivframe2.to_excel("TherapyDuration.xlsx", sheet_name='TherapyDuration', index=False);
    print("Phase IV execution ended at"+str(datetime.datetime.now().time()));
    
    input('Press ENTER to exit.');
    
    # ----------------------------END OF PROGRAM----------------------------------