import pandas as pd
import os,glob
import numpy as np
import datetime
import openpyxl

pd.set_option('display.max_columns', None)

## Change the name of variable 'file' with your excel file name
file = 'Master Books List.xlsx'
master_sheet = 'Master List'
update_sheet = 'UPDATE'

## Change the name of this sheet similar to your sheet name in excel file
df_master = pd.read_excel(file,sheet_name=master_sheet)
df_update = pd.read_excel(file,sheet_name=update_sheet)
df_update = df_update.loc[:, ~df_update.columns.str.contains('^Unnamed')]

## Save all the titles on master sheet
##master_title = (df_master.loc[:,'Title (KEY)'] + '-' + df_master.loc[:,'Author (KEY)'] + '-' + df_master.loc[:,'Date (KEY)'] + df_master.loc[:,'Page (KEY)'].astype(str) + '-' + df_master.loc[:,'URL']).tolist()
master_title = (df_master.loc[:,'Title (KEY)']+'-'+df_master.loc[:,'URL']).tolist()

## Save all the titles on update sheet
##update_title = (df_update.loc[:,'Title'] + df_update.loc[:,'Author'] + '-' + df_update.loc[:,'Date'] + '-' + df_update.loc[:,'Page'].astype(str) + '-' + df_update.loc[:,'Link']).tolist() 
update_title = (df_update.loc[:,'Title']+'-'+df_update.loc[:,'Link']).tolist() 

##master_new = df_master.loc[:,'Title (KEY)'] + '-' + df_master.loc[:,'Author (KEY)'] + '-' + df_master.loc[:,'Date (KEY)'] + df_master.loc[:,'Page (KEY)'].astype(str) + '-' + df_master.loc[:,'URL']
##update_new = (df_update.loc[:,'Title'] + df_update.loc[:,'Author'] + '-' + df_update.loc[:,'Date'] + '-' + df_update.loc[:,'Page'].astype(str) + '-' + df_update.loc[:,'Link'])
##duplicate_update = update_new[update_new.duplicated(keep='first')]
##
##np_update_new = update_new.values
##np_master_new = master_new.values
##
##rem = 0
##list_dup = []
##np_dup = duplicate_update.values
##for kl in np_dup:
##    index_dup_master = np.where(np_master_new == kl)
##    index_dup_update = np.where(np_update_new == kl)
##    if len(index_dup_master[0]) < len(index_dup_update[0]):
##        if len(index_dup_master[0]) == 0:
##            rem = 1
##            for al in index_dup_update[0]:
##                list_dup.append(al)
##        else:
##            for index_ms in index_dup_master[0]:
##                for index_up in index_dup_update[0]:
##                    if index_up == index_ms:
##                        continue
##
##            list_dup.append(index_up)
##
##print(update_new[list_dup])
##
#### Find current date and change the format to Year-Month-Day Hour:Minute:Second
##date = datetime.datetime.now()
##str_date = date.strftime('%d-%b')
##
##nw_title,nw_author,nw_date,nw_page,nw_link,time_update = [],[],[],[],[],[]
##for nw_dt in list_dup:
##    nw_title.append(df_update.loc[nw_dt,'Title'])
##    nw_author.append(df_update.loc[nw_dt,'Author'])
##    nw_date.append(df_update.loc[nw_dt,'Date'])
##    nw_page.append(df_update.loc[nw_dt,'Page'])
##    nw_link.append(df_update.loc[nw_dt,'Link'])
##    time_update.append(str_date)

## Find the difference book title between master and update sheet
b = list(set(update_title) - set(master_title))
print(len(b))
##if rem == 0:
##    print(b)
##    pass
##else:
##    b = np.array(b)
##    list_dup = list(set(update_new[list_dup]))
##    print(b)
##    for k in list_dup:
##        idx = np.where(b == k)
##        if len(idx) == 0:
##            pass
##        else:
##            b = np.delete(b,idx)
##
##    print(b)
##
##print(len(b))
##    
####if len(b) > 0:
####    numpy_b = np.array(b)
####    numpy_update = np.array(update_title)
####
####    for x in b:
####        arr_index = np.where(numpy_update == x)
####        nw_title.append(df_update.loc[arr_index[0][0],'Title'])
####        nw_author.append(df_update.loc[arr_index[0][0],'Author'])
####        nw_date.append(df_update.loc[arr_index[0][0],'Date'])
####        nw_page.append(df_update.loc[arr_index[0][0],'Page'])
####        nw_link.append(df_update.loc[arr_index[0][0],'Link'])
####        time_update.append(str_date)
####
####    ## Change all list variables to Series and convert it to DataFrame
####    title_series = pd.Series(nw_title)
####    author_series = pd.Series(nw_author)
####    date_series = pd.Series(nw_date)
####    page_series = pd.Series(nw_page)
####    link_series = pd.Series(nw_link)
####    update_series = pd.Series(time_update)
####    maj_series = pd.Series([''])
####    sub_series = pd.Series([''])
####    find_series = pd.Series([''])
####
####    frame = {'Rpt date': update_series, 'Maj Cls': maj_series, 'Sub Cls': sub_series, 'Title (KEY)': title_series, 'Author (KEY)': author_series, 'Date (KEY)':date_series, 'Page (KEY)':page_series, 'URL':link_series, 'Find ->':find_series} 
####    result = pd.DataFrame(frame)
####
####    ## Append 'result' DataFrame to 'master' sheet DataFrame
####    res_excel = df_master.append(result,ignore_index=True)
####    print(res_excel)
######    res_excel['Rpt date'] = pd.to_datetime(res_excel['Rpt date']).dt.strftime('%d-%b')
####    
####    ## save the result to the excel file
####    writer = pd.ExcelWriter('NEW.xlsx',engine='xlsxwriter')
######    res_excel.sort_values(['Title (KEY)'],ascending=True,axis=0,inplace=True,)
####    res_excel.to_excel(writer,sheet_name=master_sheet,index = False)
####    df_update.to_excel(writer,sheet_name=update_sheet,index = False)
####    writer.save()
####
####else:
####
####    print('NO DIFFERENCE')
