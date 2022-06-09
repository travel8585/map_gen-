import os
import re
import xlrd
import xlwt
from xlutils.copy import copy

book = xlrd.open_workbook('test.xls')

wb = copy(book)





sheet1 = book.sheet_names()
num_sheets =len(sheet1)

style = xlwt.easyxf('pattern: pattern solid, fore_colour ice_blue; font: bold off')
style_warning = xlwt.easyxf('pattern: pattern solid, fore_colour red; font: bold on, color yellow')


for i in range(num_sheets):
    print(i)
    ws_temp = book.sheet_by_index(i)
    wb_sheet = wb.get_sheet(i)
    num_rows = ws_temp.nrows
    num_cols = ws_temp.ncols
    print(num_rows)
    print(num_cols)
    sta_ind = 0
    end_ind = 0
    siz_ind = 0
    for j in range(num_cols):
        if ws_temp.cell_value(0,j) == 'sub start address':
            sta_ind = j
        elif ws_temp.cell_value(0,j) == 'sub end address':
            end_ind = j
        elif ws_temp.cell_value(0,j) == 'sub size':
            siz_ind = j
    if sta_ind+end_ind+siz_ind > 0:
        print(sheet1[i]+'  start:'+str(sta_ind)+'  end:'+str(end_ind)+'  size:'+str(siz_ind))
        for k in range(1,num_rows):
            if(k==1):
                sta_last = ''
                #end_last = ''
                siz_dec_last = 0
            else:
                sta_last = sta_value
                siz_dec_last = size_dec             

                
            sta_value = ws_temp.cell_value(k,sta_ind)
            end_value = ws_temp.cell_value(k,end_ind)
            siz_value = ws_temp.cell_value(k,siz_ind)

            #sta_value = wb_sheet.cell_value(k,sta_ind)
            #end_value = wb_sheet.cell_value(k,end_ind)
            #siz_value = wb_sheet.cell_value(k,siz_ind)            
            
            #print(siz_value)
            if(siz_value != ''):
                if(siz_value[-1]=='M'):
                    size_dec = int(re.sub("\D", "", siz_value))*1024*1024
                elif(siz_value[-1]=='K'):
                    size_dec = int(re.sub("\D", "", siz_value))*1024
                else:
                    size_dec = int(re.sub("\D", "", siz_value))
            else:
                size_dec = 0
            
                #print(size_dec)
            if(sta_value == '' and k>=1):
                if(sta_last != ''):
                    tmp_start = int(sta_last,16)+siz_dec_last;
                    tmp_start_hex = '{:08X}'.format(tmp_start)
                    sta_value = '0x'+tmp_start_hex[0:4]+'_'+ tmp_start_hex[4:8]
                    wb_sheet.write(k,sta_ind,sta_value,style)
                    
                    if(end_value == ''):
                        tmp_end = tmp_start+size_dec-1
                        tmp_end_hex = '{:08X}'.format(tmp_end)
                        
                        if(size_dec == 0):
                            wb_sheet.write(k,end_ind,'0x'+tmp_end_hex[0:4]+'_'+ tmp_end_hex[4:8],style_warning)
                            wb_sheet.write(k,siz_ind,'N/A!',style_warning)
                        else:
                            wb_sheet.write(k,end_ind,'0x'+tmp_end_hex[0:4]+'_'+ tmp_end_hex[4:8],style)
                            
            elif(end_value == '' and k>=1):                
                tmp_end = int(sta_value,16)+size_dec-1
                tmp_end_hex = '{:08X}'.format(tmp_end)
                wb_sheet.write(k,end_ind,'0x'+tmp_end_hex[0:4]+'_'+ tmp_end_hex[4:8],style)

    print('------------------')
    
book.release_resources()
del book
wb.save('new.xls')    
    

