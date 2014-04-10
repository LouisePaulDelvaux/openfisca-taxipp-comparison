# -*- coding: utf-8 -*-

from datetime import datetime
import math
import numpy as np
import pandas as pd
import os
# Architecture : 
# un xlsx contient des sheets qui contiennent des variables, chaque sheet ayant un vecteur de dates

def import_xls(param_name):
    path = os.path.dirname(__file__)
    path = ("P:/Legislation/Barèmes IPP/Barèmes IPP - ").encode('cp1252') 
    return pd.ExcelFile(path+ param_name +'.xlsx')

def clean_date(date):
        date = date.replace(day = 1)
        date = datetime.date(date)
        return date
        
def clean_sheet(sheet_name):
    ''' Cleaning excel sheets and creating small database'''

    sheet = xlsxfile.parse(sheet_name, index_col = None)

    # Conserver les bonnes colonnes : on drop tous les "Unnamed"
    for col in sheet.columns.values:     
        if (col[0:7] == 'Unnamed'):
            sheet = sheet.drop([col],axis = 1)
   
    def _is_var_nan(row,col):
        fusion = False
        if isinstance(sheet.iloc[row,col],float):
            fusion = math.isnan(sheet.iloc[row,col])
        return fusion
    
    # Conserver les bonnes lignes : on drop s'il y a du texte ou du NaN dans la colonne des dates
    sheet['date_renseignees'] = False
    for i in range(0,sheet.shape[0]):
        sheet.loc[i,['date_renseignees']] = isinstance(sheet.iat[i,0],unicode) | _is_var_nan(i,0)
    sheet = sheet[sheet.date_renseignees == False]
    sheet = sheet.drop(['date_renseignees'],axis = 1)
    
    # S'il y a du texte au milieu du tableau (explications par exemple) => on le transforme en NaN
    for col in range(0,sheet.shape[1]):
        for row in range(0,sheet.shape[0]):
            if isinstance(sheet.iloc[row,col],unicode):
                sheet.iat[row,col] = 'NaN'
                            
    # Gérer la suppression et la création progressive de dispositifs
    sheet.iloc[0, :] = sheet.iloc[0, :].fillna('-')
    
    # TODO: Handle currencies (Pb : on veut ne veut diviser que les montants et valeurs monétaires mais pas les taux ou conditions).
    # TODO: Utiliser les lignes supprimées du début pour en faire des lables
    # TODO: Utiliser les lignes supprimées de la fin et de la droite donner des informations sur la législation (références, notes...)
    
    try:
        sheet['date'] =[ clean_date(d) for d in  sheet['date']]
    except:
        raise "Aucune colonne date dans la feuille : ", sheet
    return sheet

def sheet_to_dic(sheet):   
    dic = {}
    sheet = clean_sheet(sheet)
    sheet.index = sheet['date']
    for var in sheet.columns.values:
        dic[var] = sheet[var]
    return dic
    
def dic_of_same_variable_names(param_name):
    xlsxfile =   import_xls(param_name)
    sheet_names = xlsxfile.sheet_names
    sheet_names = [ v for v in sheet_names if not v.startswith('Sommaire')|v.startswith('Outline') | v.startswith('Barème IGR')]
    dic={}
    for sheet in  sheet_names:
        dic[sheet]= clean_sheet(sheet)
    all_variables = np.zeros(1)
    multiple_names = []
    for sheet_name in  sheet_names:
        sheet = clean_sheet(sheet_name)
        columns =  np.delete(sheet.columns.values,0)
        all_variables = np.append(all_variables,columns)
    for i in range(0,len(all_variables)):
        var = all_variables[i]
        new_variables = np.delete(all_variables,i)
        if var in new_variables:
            multiple_names.append(str(var))
    multiple_names = list(set(multiple_names))
    dic_var_to_sheet={}
    for sheet_name in sheet_names:
        sheet = clean_sheet(sheet_name)
        columns =  np.delete(sheet.columns.values,0)
        for var in multiple_names:
            if var in columns:
                if var in dic_var_to_sheet.keys():
                    dic_var_to_sheet[var].append(sheet_name)
                else:
                    dic_var_to_sheet[var] = [sheet_name]
    return dic_var_to_sheet
   
if __name__ == '__main__':

    baremes = ['Prestations']
    
    for bareme in baremes :
        xlsxfile =   import_xls(bareme)
        test_duplicate = dic_of_same_variable_names(bareme)
        if len(test_duplicate) != 0 :
            print 'Au moins deux variables ont le même nom dans le classeur ' +  str(bareme).encode('cp1252')  + ':' 
            print  test_duplicate
            import pdb
            pdb.set_trace()
        sheet_names = xlsxfile.sheet_names
        # Retrait des onglets qu'on ne souhaite pas importer
        sheet_names = [ v for v in sheet_names if not v.startswith('Sommaire')|v.startswith('Outline') | v.startswith('Barème IGR')]
        
        mega_dic = {}
        for sheet in  sheet_names:
            mega_dic.update(sheet_to_dic(sheet))
        dateList = [ datetime.strptime( str(year) + '-' + str(month) + '-01' ,"%Y-%m-%d").date()  for year in range(1914, 2015) for month in range(1,13)]
        table = pd.DataFrame(index = dateList) 
        for k,v in mega_dic.iteritems():
            var_name = str(k)
            table[var_name] = np.nan
            table.loc[v.index.values, var_name] = v.values
        table = table.fillna(method ='pad')
        table = table.dropna(axis = 0, how = 'all')
        table.to_csv(bareme + '.csv')
        print "Voilà, la table agrégée de " + bareme + " est créée!"

#    sheet = xlsxfile.parse('majo_excep', index_col = None)

