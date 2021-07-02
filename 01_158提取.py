# -*- coding: utf-8 -*-
"""
Created on Mon Apr  6 18:32:37 2020

@author: Administrator
"""
#问题点：由于从158listing页面复制的数据中有些是12行，有些是11行，起始字段以“周对比2”，结束字段以SKU列为空值，注意是空值不是Nan
#这里采用判断一下是否是12行，12行写入某个块block，再从块中提取数据；否则从不提取数据

#update20201026:已测试通过，可以在listing数据较少的时候进行手动复制

import datetime
import pandas as pd
import os
import warnings
warnings.filterwarnings('ignore')

def get_data(path):    
    filename = path.split('\\')[-1].split('.')[0]    
    data = pd.read_excel(path)
    print('数据正在提取中，请稍后')

    block = pd.DataFrame()
    df = pd.DataFrame()
    block_others = []
    i = 0
    x = 0
    for row in range(3,len(data)):
        if data.loc[row,'周对比2'] == '周：':
            x = row      
            block_y = min(row + 15,len(data))  #此处是防止循环超过df的长度，就会报错index range out of 
            for row_2 in range(row + 1,block_y):
                y = 0
                if data.loc[row_2,'周对比2'] == '周：':  #识别第二个周对比，然后减去一行构成一个完成的block循环
                    y = row_2 - 1
                        
                if y > x:
                    block = data.ix[x:y,:]
                    block.reset_index(inplace = True)
                    
                    if len(block) == 11:  
                        df.loc[i,'站点国家_1'] = block.loc[1,'站点国家']
                        df.loc[i,'sellerSKU_1'] = block.loc[7,'SKU'].split(':')[1]
                        df.loc[i,'ASIN_1'] = block.loc[8,'SKU'].split(':')[1].lstrip()
                        df.loc[i,'中文名称_1'] = block.loc[2,'SKU']
                        df.loc[i,'节日标识'] = block.loc[4,'产品信息']

                    
                    elif len(block) == 12:
                        df.loc[i,'站点国家_1'] = block.loc[1,'站点国家']
                        df.loc[i,'sellerSKU_1'] = block.loc[7,'SKU'].split(':')[1]
                        df.loc[i,'ASIN_1'] = block.loc[8,'SKU'].split(':')[1].lstrip()
                        df.loc[i,'中文名称_1'] = block.loc[2,'SKU']
                        df.loc[i,'节日标识'] = block.loc[4,'产品信息']

                    
                    elif len(block) == 13:
                        df.loc[i,'站点国家_1'] = block.loc[1,'站点国家']
                        df.loc[i,'sellerSKU_1'] = block.loc[7,'SKU'].split(':')[1]
                        df.loc[i,'ASIN_1'] = block.loc[8,'SKU'].split(':')[1].lstrip()
                        df.loc[i,'中文名称_1'] = block.loc[2,'SKU']
                        df.loc[i,'节日标识'] = block.loc[4,'产品信息']

            
                    else:
                        print('有其他行列数据，请检查，对应行为(%d,%d)'%(x,y))
                        continue
                  
                    i += 1
                
                    break
    
    df.columns = ['渠道来源','SellSKU','ASIN','产品中文名称','节日标识']    

    origin_158_data = pd.read_excel(path)       # ,sheet_name = 'sheet1'
    writer = pd.ExcelWriter(path)
    origin_158_data.to_excel(writer,'原始158listing数据',index = False)
    data.to_excel(writer,'CPC复制数据',index = False)
    df.to_excel(writer,'已提取数据',index = False)
    writer.save()


print('请输入CPC中产品页面要提取的文件')
file_path = input('excel文件:').replace('"','')

f_path = os.path.dirname(file_path)
os.chdir(f_path)

get_data(file_path)
print('数据提取成功，路径为:%s'%f_path)

