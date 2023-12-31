from imblearn.over_sampling import SMOTE
from sklearn.preprocessing import StandardScaler, OneHotEncoder, LabelEncoder
#from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
#from sklearn.linear_model import LogisticRegression
#from sklearn.ensemble import GradientBoostingRegressor
#from sklearn.linear_model import LinearRegression
from sklearn import metrics
import random
import numpy
import pandas

def Reverse_Encode(inputframe, inputarray, numlabelname, labelname):
    categorycutoffs=GetNumericalCutoffsfromCategoricalLabels(inputframe, numlabelname, labelname);
    categorylist=[];
    for k in range(int(len(categorycutoffs)/2)):
        categorylist.append(categorycutoffs[2*k]);
    encodedframe=inputframe[[labelname, numlabelname]];
    encodedframe['Encoded_Pred']=0;
    #encodedframe=pandas.DataFrame(columns=['Encoded_Pred']);
    j=0;
    for row in inputarray:
        i=0;
        high=0;
        for element in row:
            if element > high:
                label=i;
                high=element;
            i+=1;
        encodedframe['Encoded_Pred'].iloc[[j]]=label;
        j+=1;
        print(j);
    encodedframe['Encoded_Pred'].replace(list(range(len(categorycutoffs)/2)),categorylist, inplace=True);
    return encodedframe;

def ClassImbalanceTrainTestSplit(inputframe, labelname):
    testframe=inputframe[0:0];
    trainframe=inputframe[0:0];
    if labelname in inputframe.columns:
        print('Label found.');
        ifgroup=inputframe.groupby(labelname);
        num_of_labels=0;
        min_data=len(inputframe);
        for a_label, grouped_data in ifgroup:
            if len(grouped_data) < min_data:
                min_data=len(grouped_data);
            num_of_labels+=1;
        if num_of_labels*min_data > 0.5*len(inputframe):
            data_per_label=int(0.5*len(inputframe)/num_of_labels);
        else:
            data_per_label=min_data;
        for a_label, grouped_data in ifgroup:
            locality=list(range(len(grouped_data)));
            random.shuffle(locality);
            trainframe=trainframe.append(grouped_data.iloc[locality[0:data_per_label]]);
            testframe=testframe.append(grouped_data.iloc[locality[data_per_label:]]);
    splitframelist=[trainframe, testframe];
    return splitframelist;

def ClassImbalanceTrainTestSplit70Percent(inputframe, labelname):
    testframe=inputframe[0:0];
    trainframe=inputframe[0:0];
    if labelname in inputframe.columns:
        print('Label found.');
        ifgroup=inputframe.groupby(labelname);
        num_of_labels=0;
        min_data=len(inputframe);
        for a_label, grouped_data in ifgroup:
            if len(grouped_data) < min_data:
                min_data=len(grouped_data);
            num_of_labels+=1;
        if num_of_labels*min_data > 0.5*len(inputframe):
            data_per_label=int(0.5*len(inputframe)/num_of_labels);
        else:
            data_per_label=min_data;
        for a_label, grouped_data in ifgroup:
            locality=list(range(len(grouped_data)));
            random.shuffle(locality);
            trainframe=trainframe.append(grouped_data.iloc[locality[0:int(data_per_label*0.7)]]);
            testframe=testframe.append(grouped_data.iloc[locality[int(data_per_label*0.7):]]);
    splitframelist=[trainframe, testframe];
    return splitframelist;

def SetPercentileCutoff(dataframe, numlabelname, percentile):
    dataframe=dataframe.sort_values(by=numlabelname, ascending=True);
    total=len(dataframe);
    dataframe[str(int(percentile))+' Percentile Cutoff+']=0;
    if 0 < int(percentile) < 100:
        measure=0;
        for index, row in dataframe.iterrows():
            measure+=1;
            if float(measure/total) > float(percentile)/100:
                dataframe[str(int(percentile))+' Percentile Cutoff+'][measure-1:]=1;
                break;
    else:
        print('Input not in range of percentile.');
    return dataframe;

def SetAbsoluteCutoff(dataframe, numlabelname, absolute):
    dataframe[str(int(absolute))+' Absolute Cutoff+']=0;
    dataframe=dataframe.sort_values(by=numlabelname, ascending=True);
    i=0;
    for index, row in dataframe.iterrows():
        if dataframe[numlabelname].iloc[[i]].item() >= int(absolute):
            dataframe[str(int(absolute))+' Absolute Cutoff+'][i:]=1;
            break;
            #dataframe[str(int(absolute))+' Absolute Cutoff+'].iloc[[i]]=1;
        i+=1;
    return dataframe;

def GetNumericalCutoffsfromCategoricalLabels(dataframe, numlabelname, labelname):
    dataframe=dataframe.sort_values(by=numlabelname, ascending=True);
    numericalcutoffs=[];
    labelgroup=dataframe.groupby(labelname);
    for labels, examples in labelgroup:
        numericalcutoffs.append(labels);
        numericalcutoffs.append(float(examples[numlabelname].iloc[[0]].item()));
    return numericalcutoffs;

def SetCategoricalLabelsfromNumericalLabels(dataframe, prednumlabelname, predlabelname, numericalcutoffs):
    dataframe=dataframe.sort_values(by=prednumlabelname, ascending=True);
    i=-1;
    j=k=0;
    for index, row in dataframe.iterrows():
        if numericalcutoffs[i] <= row[prednumlabelname] < numericalcutoffs[i-2]:
            j+=1;
            continue; #iterate
        else:
            if -i < len(numericalcutoffs)-3:
                dataframe[predlabelname][k:]=numericalcutoffs[i-1];
                k=j;
                i-=2;
            else:
                dataframe[predlabelname][k:]=numericalcutoffs[i-1];
                dataframe[predlabelname][j:]=numericalcutoffs[i-3];
                break;
        j+=1;
    '''for i in range(int(len(dataframe))):
        for j in range(int(len(numericalcutoffs)/2)):
            if j==0 and dataframe[prednumlabelname].iloc[[i]].item() >= numericalcutoffs[1]:
                dataframe[predlabelname].iloc[[i]]=numericalcutoffs[0];
            else:
                if numericalcutoffs[2*j-1] <= dataframe[prednumlabelname].iloc[[i]].item() <= numericalcutoffs[2*j+1]:
                    dataframe[predlabelname].iloc[[i]]=numericalcutoffs[2*j];'''
    predlabelgroup=dataframe.groupby(predlabelname);
    num=0;
    for label, examples in predlabelgroup:
        num+=1;
    print(str(num)+" predicted categorical labels found.");
    return dataframe;

def GeneratePerformanceMatrix(resultframe, labelname, predlabelname):
    labelgroup=resultframe.groupby(labelname);
    predlabelgroup=resultframe.groupby(predlabelname);
    labellist=[];
    predlabellist=[];
    for label, examples_1 in labelgroup:
        labellist.append(label);
    labellist.append('Accuracy');
    for predlabel, examples_2 in predlabelgroup:
        predlabellist.append(predlabel);
    performancematrix=pandas.DataFrame(index=labellist, columns=predlabellist);
    labellist.remove('Accuracy');
    for label, examples in labelgroup:
        for predlabel in predlabellist:
            numincommon=0;
            sublabelgroup=examples.groupby(predlabelname);
            for subpredlabel, common_examples in sublabelgroup:
                if subpredlabel == predlabel:
                    numincommon=len(common_examples);
            performancematrix[predlabel][label]=numincommon;
    performancematrix['Total']=0;
    nettotal=0;
    for label in labellist:
        total=0
        for predlabel in predlabellist:
            total+=performancematrix[predlabel][label];
        performancematrix['Total'][label]=total;
        nettotal+=total;
    accuratepredictions=0;
    for predlabel in predlabellist:
        performancematrix[predlabel]['Accuracy']=round(100*performancematrix[predlabel][predlabel]/performancematrix['Total'][predlabel], 2);
        accuratepredictions+=performancematrix[predlabel][predlabel];
    performancematrix['Total']['Accuracy']=round(100*accuratepredictions/nettotal, 2);
    return performancematrix;

def GenerateKPI(resultframe, labelname, predlabelname):
    labelgroup=resultframe.groupby(labelname);
    predlabelgroup=resultframe.groupby(predlabelname);
    labellist=[];
    predlabellist=[];
    for label, examples_1 in labelgroup:
        labellist.append(label);
    labellist.append('Accuracy');
    for predlabel, examples_2 in predlabelgroup:
        predlabellist.append(predlabel);
    performancematrix=pandas.DataFrame(index=labellist, columns=predlabellist);
    labellist.remove('Accuracy');
    for label, examples in labelgroup:
        for predlabel in predlabellist:
            numincommon=0;
            sublabelgroup=examples.groupby(predlabelname);
            for subpredlabel, common_examples in sublabelgroup:
                if subpredlabel == predlabel:
                    numincommon=len(common_examples);
            performancematrix[predlabel][label]=numincommon;
    performancematrix['Total']=0;
    nettotal=0;
    for label in labellist:
        total=0
        for predlabel in predlabellist:
            total+=performancematrix[predlabel][label];
        performancematrix['Total'][label]=total;
        nettotal+=total;
    accuratepredictions=0;
    for predlabel in predlabellist:
        performancematrix[predlabel]['Accuracy']=round(100*performancematrix[predlabel][predlabel]/performancematrix['Total'][predlabel], 2);
        accuratepredictions+=performancematrix[predlabel][predlabel];
    performancematrix['Total']['Accuracy']=round(100*accuratepredictions/nettotal, 2);
    performancematrix.loc['Total Pred']='-';
    for predlabel in predlabellist:
        performancematrix[predlabel]['Total Pred']=0;
        for label in labellist:
            performancematrix[predlabel]['Total Pred']+=int(performancematrix[predlabel][label]);
    performancematrix.loc['_break_']='---';
    performancematrix.loc['DR']='-';
    for predlabel in predlabellist:
        performancematrix.loc['_'+predlabel+'_']='-';
        for label in labellist:
            if label == predlabel:
                performancematrix[label]['_'+predlabel+'_']=round(100*performancematrix[predlabel][predlabel]/performancematrix['Total'][predlabel], 2);
    performancematrix.loc['__break__']='---';
    performancematrix.loc['LQ']='-';
    for predlabel in predlabellist:
        performancematrix.loc['__'+predlabel+'__']='-';
        for label in labellist:
            if label == predlabel:
                performancematrix[label]['__'+predlabel+'__']=round(100*performancematrix[predlabel][predlabel]/performancematrix[predlabel]['Total Pred'], 2);
    return performancematrix;

if __name__ == '__main__':
    
    onehotenc=OneHotEncoder(handle_unknown='ignore');
    labelenc=LabelEncoder();
    option=False;
    while not option:
        try:
            excelfilename=input("Enter the name of the excel file to be imported(w/ext.):");
            excelsheetname=input("Enter the sheet name in the excel file of interest:");
            dataframe=pandas.read_excel(excelfilename, sheet_name=excelsheetname);
            print("Data file successfully imported.");
            for attribute in dataframe.columns:
                dataframe[attribute]=dataframe[attribute].apply(lambda x: 0 if x=='' else x);
                dataframe[attribute]=dataframe[attribute].apply(lambda x: 0 if x=='_NULL' else x);
                dataframe.fillna(0, inplace=True);
            print("_NULL and [] Replaced with 0.");
            option=True;
        except AttributeError:
            print("There's no item with that code");
        except KeyError:
            print("Bad parameter name");
        except:
            print("Unknown error");
    option1=option2=option3=False;
    while not option1 or not option2 or not option3:
        try:
            idname=input("Enter the ID attribute name:");
            if idname in dataframe.columns:
                print("ID found.");
                option1=True;
            else:
                print("Label not found. Please Re-enter.");
                option1=False;
            labelname=input("Enter the categorical label name:");
            if labelname in dataframe.columns:
                labelgroup=dataframe.groupby(labelname);
                i=0;
                for label, examples in labelgroup:
                    i+=1;
                print(str(i)+" categorical labels found.");
                option2=True;
            else:
                print("Label not found. Please Re-enter.");
                option2=False;
            numlabelname=input("Enter the numerical label name:");
            if numlabelname in dataframe.columns:
                print("Numerical Label found.");
                option3=True;
            else:
                print("Label not found. Please Re-enter.");
                option3=False;
            if idname == numlabelname or numlabelname == labelname or labelname == idname:
                print("User entered identical columns for ID, Numerical label or Categorical label. Please Re-enter.")
                option1=False;
        except AttributeError:
            print("There's no item with that code");
        except KeyError:
            print("Bad parameter name");
        except:
            print("Unknown error");
    option=False;
    while not option:
        try:
            percentvarnum=input('Enter the number of percentile cutoff based variables to be generated:');
            percentilecutoffs=[];
            if int(percentvarnum) > 0:
                for i in range(int(percentvarnum)):
                    temp=input("Enter the "+str(i+1)+"th percentile cutoff:");
                    percentilecutoffs.append(temp);
                print("Generating the percentile cutoff based variables...");
                for i in range(int(percentvarnum)):
                    dataframe=SetPercentileCutoff(dataframe, numlabelname, percentilecutoffs[i]);
                print("...done.");
                option=True;
            elif int(percentvarnum) == 0:
                option=True;
                print("No percentile based variables created.");
            else:
                print("Invalid Input.");
        except AttributeError:
            print("There's no item with that code");
        except KeyError:
            print("Bad parameter name");
        except:
            print("Unknown error");
    option=False;
    while not option:
        try:
            absolutevarnum=input('Enter the number of absolute cutoff based variables to be generated:');
            absolutecutoffs=[];
            if int(absolutevarnum) > 0:
                for i in range(int(absolutevarnum)):
                    temp=input("Enter the "+str(i+1)+"th absolute cutoff:");
                    absolutecutoffs.append(temp);
                for i in range(int(absolutevarnum)):
                    dataframe=SetAbsoluteCutoff(dataframe, numlabelname, absolutecutoffs[i]);
                print("...done.");
                option=True;
            elif int(absolutevarnum) == 0:
                print("No absolute based variables created.");
                option=True;
            else:
                print("Invalid Input.");
        except AttributeError:
            print("There's no item with that code");
        except KeyError:
            print("Bad parameter name");
        except:
            print("Unknown error");
    retry=True;
    while retry:
        attributearray=dataframe.columns;
        attributelist=list(attributearray);
        attributelist.remove(numlabelname);
        attributelist.remove(labelname);
        attributelist.remove(idname);
        attributelistcopy=attributelist.copy();
        print("Select the attributes for Random Forest model(Y/N and ENTER):");
        numofselectedattributes=0;
        for attribute in attributelistcopy:
            print(attribute);
            option=False;
            while not option:
                choice=input();
                if choice == 'Y' or choice == 'y':
                    if dataframe[attribute].dtype not in ['int16', 'int32', 'int64', 'float16', 'float32', 'float64']:
                        dataframe[attribute] = labelenc.fit_transform(dataframe[attribute].astype(str));
                        #encodingarray=dataframe[attribute].values;
                        #encodingarray[:]=labelenc.fit_transform(encodingarray[:]);
                        print("...labels encoded.")
                    print("...selected.");
                    attributelist.remove(attribute);
                    attributelist.append(attribute);
                    numofselectedattributes+=1;
                    option=True;
                elif choice =='N' or choice == 'n':
                    print("...not unselected.");
                    option=True;
                else:
                    print("Not an option...");
        attributelist.append(numlabelname);
        attributelist.append(labelname);
        attributelist.append(idname);
        dataframe=dataframe[attributelist];
        sc=StandardScaler();
        regressor=RandomForestRegressor(n_estimators=100, max_depth=100, random_state=78);
        #regressor=LogisticRegression();
        #regressor=LinearRegression();
        #regressor=GradientBoostingRegressor();
        option2=False;
        while not option2:
            choice2=input("Should the data be balanced synthetically by adding extra samples to minority class(Y/N and ENTER)?");
            if choice2 == 'Y' or choice2 == 'y':
                synthesis=True;
                option2=True;
            elif choice2 == 'N' or choice2 =='n':
                synthesis=False;
                splitframelist=ClassImbalanceTrainTestSplit70Percent(dataframe, labelname);
                trainframe=splitframelist[0];
                testframe=splitframelist[1];
                print("Training data is balanced by collecting 70% of minority class's examples from every class.");
                option2=True;
        option=False;
        while not option:
            choice=input("Do you wish to only use the categorical variable(Y/N and ENTER)?");
            if choice == 'N' or choice =='n':
                if not synthesis:
                    X_train=trainframe.iloc[:, len(attributelist)-int(numofselectedattributes)-3:len(attributelist)-3].values;
                    X_test=testframe.iloc[:, len(attributelist)-int(numofselectedattributes)-3:len(attributelist)-3].values;
                    y_train=trainframe.iloc[:, len(attributelist)-3].values;
                    y_test=testframe.iloc[:, len(attributelist)-3].values;
                    # No test and train data created
                    '''X = dataframe.iloc[:, len(attributelist)-int(numofselectedattributes)-3:len(attributelist)-3].values;  
                    y = dataframe.iloc[:, len(attributelist)-3].values;
                    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.05, random_state=0);'''
                    X_train_2 = sc.fit_transform(X_train);  
                    X_test_2 = sc.transform(X_test);
                    regressor.fit(X_train_2, y_train);
                    X_total=numpy.concatenate([X_train, X_test]);
                    X_total_db=numpy.concatenate([X_train_2, X_test_2]);
                    y_total=numpy.concatenate([y_train, y_test]);
                else:
                    sm=SMOTE(random_state=2);
                    X_total=dataframe.iloc[:, len(attributelist)-int(numofselectedattributes)-3:len(attributelist)-3].values;
                    y_total=dataframe.iloc[:, len(attributelist)-3].values;
                    X_total_2=sc.fit_transform(X_total);
                    X_synth, y_synth = sm.fit_sample(X_total_2, y_total.ravel());
                    regressor.fit(X_synth, y_synth);
                    X_total_db=X_total_2;
                y_pred = regressor.predict(X_total_db);
                # Original Code
                '''X_2=sc.fit_transform(X);
                regressor.fit(X_2, y);
                y_pred=regressor.predict(X_2);'''
                print('Mean Absolute Error:', metrics.mean_absolute_error(y_total, y_pred));
                print('Mean Squared Error:', metrics.mean_squared_error(y_total, y_pred));
                print('Root Mean Squared Error:', numpy.sqrt(metrics.mean_squared_error(y_total, y_pred)));
                # Old code
                y_total_frame=pandas.DataFrame(y_total);
                y_pred_frame=pandas.DataFrame(y_pred);
                dataset=pandas.DataFrame(columns=['Y_TOTAL', 'Y_PRED']);
                dataset['Y_TOTAL']=y_total_frame[0];
                dataset['Y_PRED']=y_pred_frame[0];
                #dataset.to_excel('ResultRandomForestRegr.xlsx', sheet_name='Result');
                #y_pred_frame=pandas.DataFrame(y_pred);
                dataframe=dataframe.sort_values(by=numlabelname, ascending=True);
                dataset=dataset.sort_values(by='Y_TOTAL', ascending=True);
                if synthesis:
                    dataframe2=dataframe;
                else:
                    dataframe2=pandas.concat([trainframe, testframe]);
                dataframe2['Predicted '+numlabelname]=dataset['Y_PRED'];
                dataframe2['Predicted '+labelname]='';
                dataframe3=SetCategoricalLabelsfromNumericalLabels(dataframe2, \
                                                                   'Predicted '+numlabelname, \
                                                                   'Predicted '+labelname, \
                                                                   GetNumericalCutoffsfromCategoricalLabels(dataframe2, numlabelname, labelname));
                performancematrix=GeneratePerformanceMatrix(dataframe3, labelname, 'Predicted '+labelname);
                print(performancematrix);
                print("The model's accuracy is "+str(performancematrix['Total']['Accuracy'])+"%.");
                dataframe3.to_excel("[Pred]"+excelfilename, sheet_name=excelsheetname);
                performancematrix.to_excel("[Perf]"+excelfilename, sheet_name='Performance Matrix');
                option=True;
            elif choice == 'Y' or choice == 'y':
                if not synthesis:
                    X_train=trainframe.iloc[:, len(attributelist)-int(numofselectedattributes)-3:len(attributelist)-3].values;
                    X_test=testframe.iloc[:, len(attributelist)-int(numofselectedattributes)-3:len(attributelist)-3].values;
                    y_train=trainframe.iloc[:, len(attributelist)-2].values;
                    y_test=testframe.iloc[:, len(attributelist)-2].values;
                    y_train=y_train.reshape(-1, 1);
                    y_test=y_test.reshape(-1, 1);
                    y_train_2=labelenc.fit_transform(y_train);
                    y_test_2=labelenc.transform(y_test);
                    y_train_2=y_train_2.reshape(-1, 1);
                    y_test_2=y_test_2.reshape(-1, 1);
                    y_train_3=onehotenc.fit_transform(y_train_2).toarray();
                    y_test_3=onehotenc.transform(y_test_2).toarray();
                    X_train_2 = sc.fit_transform(X_train);  
                    X_test_2 = sc.transform(X_test);
                    regressor.fit(X_train_2, y_train_3);
                    X_total_db=numpy.concatenate([X_train_2, X_test_2]);
                    y_total=numpy.concatenate([y_train_3, y_test_3]);
                else:
                    sm=SMOTE(random_state=2);
                    X_total=dataframe.iloc[:, len(attributelist)-int(numofselectedattributes)-3:len(attributelist)-3].values;
                    y_total=dataframe.iloc[:, len(attributelist)-2].values;
                    y_total=y_total.reshape(-1, 1);
                    y_total_2=labelenc.fit_transform(y_total);
                    X_total_2=sc.fit_transform(X_total);
                    X_synth, y_synth = sm.fit_sample(X_total_2, y_total_2.ravel());
                    y_synth=y_synth.reshape(-1, 1);
                    y_synth_2=onehotenc.fit_transform(y_synth).toarray();
                    regressor.fit(X_synth, y_synth_2);
                    X_total_db=X_total_2
                y_pred=regressor.predict(X_total_db);
                #y_pred_2=onehotenc.inverse_transform(y_pred);
                y_pred_2=y_pred.dot(onehotenc.active_features_).astype(int);
                #y_pred_3=labelenc.inverse_transform(y_pred_2);
                labelenc_map_dict=dict(zip(labelenc.classes_, labelenc.transform(labelenc.classes_)));
                inverted_map_dict = dict(map(reversed, labelenc_map_dict.items()));
                y_pred_3=numpy.array([inverted_map_dict[letter] for letter in y_pred_2]);
                if synthesis:
                    dataframe2=dataframe;
                else:
                    dataframe2=pandas.concat([trainframe, testframe]);
                dataframe2['Predicted '+labelname]=y_pred_3;
                performancematrix=GenerateKPI(dataframe2, labelname, 'Predicted '+labelname);
                print(performancematrix);
                print("The model's accuracy is "+str(performancematrix['Total']['Accuracy'])+"%.");
                dataframe2.to_excel("[Pred]"+excelfilename, sheet_name=excelsheetname);
                performancematrix.to_excel("[Perf]"+excelfilename, sheet_name='Performance Matrix');
                option=True;
            else:
                print("Not an option...");
        option=False;
        while not option:
            choice=input("Do you wish to retry with different selected attributes?(Y/N and ENTER):");
            if choice == 'Y' or choice == 'y':
                print("...restarting the program.");
                option=True;
            elif choice =='N' or choice == 'n':
                retry=False;
                option=True;
            else:
                print("Not an option...");
    input("Press ENTER to exit.");