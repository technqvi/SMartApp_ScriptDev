#!/usr/bin/env python
# coding: utf-8

# In[1]:


import psycopg2
import psycopg2.extras as extras
from psycopg2.extensions import AsIs
import numpy as np
import pandas as pd
import os
from datetime import datetime
from dataclasses import dataclass
from dotenv import dotenv_values


run_py=True


# In[ ]:


inventory_file="../Inventory_YIT_BlockChain2.xlsx"
error_file="../Error_Inventory.xlsx"

if run_py:
    inventory_name=input("Enter Inventory Excel Fild(ex. Inventory_XYZ.xlsx) :")
    inventory_file=os.path.join("..",inventory_name)
    if  os.path.exists(inventory_file) : 
        
     xname,xtype=os.path.splitext(inventory_name)
     error_file=os.path.join("..",f"Error_{xname}{xtype}")
        
        
     print(f"Inventory Path: {os.path.abspath(inventory_file)}")
     print(f"Error Path (if errors): {os.path.abspath(error_file)} ")
     
     y_n = input(f"Confirm import inventory ,please press y:")
     if y_n!='y':
        exit()
    
    else:
     print(f"Not found excel file  {os.path.abspath(inventory_file)}")   
     exit() 


# In[139]:


# eror case
#inventory_file="inventory_import/Inventory_All-Error.xlsx"
#inventory_file="inventory_import/Inventory_Incomplete.xlsx"

# complter case
#inventory_file="inventory_import/Inventory_Master.xlsx"
#inventory_file="inventory_import/Inventory_MEA_060622_2025.xlsx"

# get from database or file
inventory_schema='InventoryExport_Schema.xlsx'

no_records=1

is_inccorect_data=False
list_error=[]


# In[ ]:





# In[140]:


def isnan(value):
  try:
      import math
      return math.isnan(float(value))
  except:
      return False


# In[141]:


@dataclass
class MappingData:
    disp_name: str
    pk_id: int
    search_name: str
    sql_cmd: str
    params: dict

# remove it in production
def get_postgres_conn():
 try:
  config = dotenv_values(dotenv_path='.env')  
  conn = psycopg2.connect(
         database=config['DATABASES_NAME'], user=config['DATABASES_USER'],
      password=config['DATABASES_PASSWORD'], host=config['DATABASES_HOST'],
     )
  return conn

 except Exception as error:
  print(error)      
  raise error
    
def list_data(sql,params,connection):
 df=None   

 with connection.cursor() as cursor:
    # print(sql)
    # print(params)    
    
    if params is None:
       cursor.execute(sql)
    else:
       cursor.execute(sql,params) 
    
    columns = [col[0] for col in cursor.description]
    dataList = [dict(zip(columns, row)) for row in cursor.fetchall()]
    df = pd.DataFrame(data=dataList) 
 return df 


# In[142]:


print("========================================================================")
print("Load inventory schema mapping between excel report and inventory table")
df_schema=pd.read_excel(inventory_schema)
df_schema=df_schema.sort_values(by=['IsPK','IsNULL','DisplayName'],ascending=False)
#print(df_schema)
print(df_schema[['DisplayName','ColumnName']])

metaDF_pk=df_schema.query('IsPK==1').set_index('ColumnName')
metaDF_string=df_schema.query('IsString==1').set_index('ColumnName')
metaDF_notNull=df_schema.query('IsNULL==0').set_index('ColumnName')


# print(metaDF_pk)
# print(metaDF_string)
#print(metaDF_notNull)


# In[143]:


main_cols_error_file=['serial_number','quantity','project_id'
                      ,'customer_warranty_start','customer_warranty_end','customer_sla_id' \
                      ,'yit_warranty_start','yit_warranty_end','yit_sla_id'
                      ,'product_warranty_start','product_warranty_end','product_sla_id'
                      ,'product_type_id','brand_id','model_id','branch_id','datacenter_id'\
                      ,'customer_support_id','customer_pm_support_id','cm_serviceteam_id'
                      ,'pm_serviceteam_id','cm_serviceteam_id','product_support_id','function_id'
                      ,'install_date','eos_date'
                     ]

#main_cols_error_file
second_cols_error_file=  [ x for x in df_schema['ColumnName'].tolist() if x not in  main_cols_error_file ]
#second_cols_error_file
main_cols_error_file.extend(second_cols_error_file)
#main_cols_error_file


# In[ ]:





# # Load Excel Inventory and Check Data Format and Null Value

# In[144]:


print("========================================================================")
print("Load excel report inventory to import")

try:
    df_new=pd.read_excel(inventory_file)
    df_new=df_new[df_schema['DisplayName'].tolist()]
    print(df_new.head(10))
#print(df_new.info())

except Exception as ex:
   error=f"Some columns in excel doestn't match exactly with inventory schema\n {str(ex)}"
   raise Exception(error)

# # check# aolumne
# a=list(df_new.columns)
# b=list(df_schema['DisplayName'])
# a_diff_b= list ( set(a)^set(b) )


# In[145]:


print("Check Data Format and Null Value")
print("========================================================================")


# In[146]:


# check# no-recourd
if df_new.shape[0]<=no_records :
    list_error.append(f"Number of inventory is less than {no_records}")
else:
    print(f"Number of inventory  is more than {no_records}")


# In[147]:


# check# null
null_cols=df_new[list(metaDF_notNull['DisplayName'])].isnull().sum()
null_cols=null_cols[null_cols>0]
if not null_cols.empty:
   list_error.append("found empty value in some columns in excel file : \n"+null_cols[null_cols>0].to_string())
else:
   print("there is no null value in required columns")


# In[148]:


# check# convert datatime 
# df_new['Install Date']=pd.to_datetime (df_new['Install Date'],format='%d %b %Y')
# df_new['EOS Date']=pd.to_datetime (df_new['EOS Date'],format='%d %b %Y')
# don't convert to datetime straightforwardli in order to advoid havong some NaT value for None datetime

def convert_datetime_string_format(item):
  if isnan(item)==False:
    try:
      d_date =datetime.strptime(item, '%d %b %y')
      d_str =d_date.strftime('%Y-%m-%d')
      return   d_str
    except Exception as ex:
       raise ex    
  return item  
try:

    df_new['Cust Warranty Start']=pd.to_datetime (df_new['Cust Warranty Start'],format='%d %b %Y')
    df_new['Cust Warranty End']=pd.to_datetime (df_new['Cust Warranty End'],format='%d %b %Y')
    df_new['Yit Warranty Start']=pd.to_datetime (df_new['Yit Warranty Start'],format='%d %b %Y')
    df_new['Yit Warranty End']=pd.to_datetime (df_new['Yit Warranty End'],format='%d %b %Y')
    df_new['Product Warranty Start']=pd.to_datetime (df_new['Product Warranty Start'],format='%d %b %Y')
    df_new['Product Warranty End']=pd.to_datetime (df_new['Product Warranty End'],format='%d %b %Y')
    
    df_new['Install Date']=df_new['Install Date'].apply(convert_datetime_string_format)
    df_new['EOS Date']=df_new['EOS Date'].apply(convert_datetime_string_format)

except Exception as ex:
   list_error.append("Wrong DateFormat : \n"+str(ex))


# In[149]:


# check# customer and product nae
try:
    df_new['Customer Support']=df_new['Customer Support'].apply( lambda x : (x.strip().split('|')[0]).strip() if (isnan(x)==False) else np.NaN )
    df_new['Customer PM Support']=df_new['Customer PM Support'].apply( lambda x : (x.strip().split('|')[0]).strip() if (isnan(x)==False) else np.NaN )
    df_new['Product Support']=df_new['Product Support'].apply( lambda x : (x.strip().split('|')[0]).strip() if (isnan(x)==False) else np.NaN )
except Exception as ex:
   list_error.append(str(ex))


# In[150]:


# check# numberic value
try:
    df_new['Storage Capacity']=df_new['Storage Capacity'].astype(float)
    df_new['QTY']=pd.to_numeric(df_new['QTY'])
except Exception as ex:
    list_error.append(str(ex))


# In[151]:


if (len(list_error)>0):
    print("Found some errors as folows")
    for i in range(len(list_error)):
     print(f"{i+1} - {list_error[i]}")
    raise Exception(f'error as the the following above, check error in {inventory_file}')

# return error report


# In[152]:


# thrown error to show


# # Starting Point to import excel to databae

# In[153]:


print("Data is ready to import")
print(df_new.info())
df_new


# In[154]:


print("Load Extract Transform and Import to Database")
print("Findk pk from name of all master table")
print("========================================================================")


# In[155]:



def get_pk_id(item,meta,df_filter):
    
  value_name=item[meta.disp_name]
 
  if isnan(value_name)==False:
    if  type(value_name)==str:
      value_name=value_name.strip()  
    
    x=df_filter.query(f'{meta.search_name}==@value_name')
    
    if len(x.index)==1:
        return x.iloc[0]['id']
    else:
        return np.nan    
        #return None
  else:
        #return None
        return value_name
    


# In[156]:


def find_pk(df_temp,pk_id,search,px,sql):
    try:

        info_x=metaDF_pk.loc[pk_id,:]
        disp_name=info_x['DisplayName']
        
        list_pkID_toQuery=tuple(df_temp[disp_name].dropna().unique())
        print(f"{pk_id} of {list_pkID_toQuery}")
        
        if len(list_pkID_toQuery)==0:
            df_temp[pk_id]=None
            return df_temp
        else:
 
            meta= MappingData( disp_name=disp_name,pk_id=pk_id,search_name=search,sql_cmd=sql,params={px:list_pkID_toQuery} )

            df_filter=list_data(meta.sql_cmd,meta.params,get_postgres_conn())
            if df_filter.empty==False:
             print("Found pk id as the belows")   
             print(df_filter)
             df_temp[meta.pk_id]=df_temp.apply(get_pk_id,axis=1,args=(meta,df_filter))
            else:
             print(f"No found any pk id along with {list_pkID_toQuery}")   
             df_temp[meta.pk_id]=np.nan
                

            print("Extract PK Id from "+disp_name)
            print("==============Found PK============")
            print(df_temp[df_new[pk_id].notnull()][[pk_id,disp_name]] .head(10))
            print("==============NotFound PK============")
            print(df_temp[df_new[pk_id].isnull()][[pk_id,disp_name]])
            
        return df_temp

    except Exception as ex:
        print( ex)


# # Mapping PK

# In[157]:


s_name='enq_id'
s_param='enq_param'


#df_new=
find_pk(df_new,'project_id',s_name, s_param,f""" select {s_name} ,id from app_project where {s_name} in %({s_param})s """ )


# In[158]:


s_name='productype_name'
s_param='productype_param'
#df_new=
find_pk(df_new,'product_type_id',s_name, s_param,                f""" select {s_name} ,id from app_product_type where {s_name} in %({s_param})s """ )


# In[159]:


s_name='customer_name'
s_param='cust_param'
df_new=find_pk(df_new,'customer_support_id',s_name, s_param,f""" select {s_name} ,id from app_customer where {s_name} in %({s_param})s """ )
df_new=find_pk(df_new,'customer_pm_support_id',s_name, s_param,f""" select {s_name} ,id from app_customer where {s_name} in %({s_param})s """ )


# In[160]:


s_name='product_name'
s_param='prod_param'
df_new=find_pk(df_new,'product_support_id',s_name, s_param,f""" select {s_name} ,id from app_product where {s_name} in %({s_param})s """ )


# In[161]:


s_name='brand_name'
s_param='brand_param'
df_new=find_pk(df_new,'brand_id',s_name, s_param,                f""" select {s_name} ,id from app_brand where {s_name} in %({s_param})s """ )


# In[162]:


s_name='model_name'
s_param='model_param'
df_new=find_pk(df_new,'model_id',s_name, s_param,                f""" select {s_name} ,id from app_model where {s_name} in %({s_param})s """ )


# In[163]:


s_name='datacenter_name'
s_param='datacenter_param'
df_new=find_pk(df_new,'datacenter_id',s_name, s_param,                f""" select {s_name} ,id from app_datacenter where {s_name} in %({s_param})s """ )


# In[164]:


s_name='branch_name'
s_param='branch_param'
df_new=find_pk(df_new,'branch_id',s_name, s_param,                f""" select {s_name} ,id from app_branch where {s_name} in %({s_param})s """ )


# In[165]:


s_name='function_name'
s_param='function_param'
df_new=find_pk(df_new,'function_id',s_name, s_param,                f""" select {s_name} ,id from app_function where {s_name} in %({s_param})s """ )


# In[166]:


s_name='function_name'
s_param='function_param'
df_new=find_pk(df_new,'function_id',s_name, s_param,                f""" select {s_name} ,id from app_function where {s_name} in %({s_param})s """ )


# In[167]:


s_name='service_team_name'
s_param='service_param'
s_sql= f""" select {s_name} ,id from app_serviceteam where {s_name} in %({s_param})s """
df_new=find_pk(df_new,'cm_serviceteam_id',s_name, s_param, s_sql)
df_new=find_pk(df_new,'pm_serviceteam_id',s_name, s_param, s_sql )


# In[168]:


s_name='sla_name'
s_param='sla_param'
s_sql= f""" select {s_name} ,id from app_sla where {s_name} in %({s_param})s """

df_new=find_pk(df_new,'customer_sla_id',s_name, s_param, s_sql)
df_new=find_pk(df_new,'yit_sla_id',s_name, s_param, s_sql )
df_new=find_pk(df_new,'product_sla_id',s_name, s_param, s_sql )


# # Get data ready for saving into database

# In[169]:


print("Mapping columns to insert into Databse")
not_pk_cols= list(set(df_schema['DisplayName'].tolist()) - set(metaDF_pk['DisplayName'].tolist()))
#not_pk_cols
#print(disp_to_col)

metaDF_notPKCols=df_schema.query('IsPK==0')
#metaDF_notPKCols

disp_to_col_notPKCols=  dict( zip( metaDF_notPKCols['DisplayName'].tolist(), metaDF_notPKCols['ColumnName'].tolist()))
#disp_to_col_notPKCols

df_new=df_new.rename(columns= disp_to_col_notPKCols)
# df_new=df_new.where(pd.notnull(df_new), None)

final_cols_to_db=metaDF_notPKCols['ColumnName'].tolist()+metaDF_pk.index.tolist()
df_new=df_new[final_cols_to_db]

df_new['is_dummy']=False
df_new['updated_at']=datetime.now()



df_cols=df_new.columns.tolist()
sql_cols_schemm="SELECT column_name FROM information_schema.columns WHERE  table_name = 'app_inventory'"
listCols_InventoryTable= list_data(sql_cols_schemm ,None,get_postgres_conn())
table_cols=listCols_InventoryTable['column_name'].tolist()

diff_cols=list(set(table_cols) -set(df_cols))
print(diff_cols)
if len(diff_cols)==1 : # except id
  print("Getting Ready to database")

print(f"{len(df_new.index)} items are about to import to database.")
print("=======================Create tempID for filtering unqualified data===============================")

df_new=df_new.reset_index(drop=False)
df_new=df_new.rename(columns={'index':'temp_id'})

df_new.info()
df_new
#df_new.to_excel('new_inventory.xlsx',index=False)


# # Check correct Data

# # Check Dupplicate Row 

# In[170]:


print("Check dupplicated record in excel file")
chekc_dup_cols=df_new.columns.tolist()
chekc_dup_cols.remove('temp_id')

df_duplicatedRows = df_new[df_new.duplicated(subset=chekc_dup_cols,keep='first')][main_cols_error_file]

if len(df_duplicatedRows.index)>0:
 is_inccorect_data=True

df_duplicatedRows


# In[171]:


df_new=df_new.drop_duplicates(subset=chekc_dup_cols,keep='first')
#df_new


# # Check  not found some pk_id values

# In[172]:


print("check  not found pk_id")
metaDF_pk_not_null=metaDF_pk=df_schema.query('IsPK==1 and IsNULL==0').set_index('ColumnName')
pkNull_df = df_new[df_new[list(metaDF_pk_not_null.index)].isnull().any(1)]


# In[173]:


pkNullCols_sr=df_new[metaDF_pk_not_null.index.tolist()].isnull().sum()
pkNullCols_sr=pkNullCols_sr[pkNullCols_sr>0]
pkNullCols_sr


# In[174]:


if len(pkNull_df.index)>0:
 is_inccorect_data=True
 df_new=df_new.drop(pkNull_df['temp_id'].tolist(), axis=0)
 pkNull_df=pkNull_df[main_cols_error_file]
   
pkNull_df 
#df_new


# # Check existing row base on ENQ ID,productType ,serial

# In[175]:


print("Check existing row base on ENQ ID,productType ,serial")
def is_existing_row(sql,params,connection):
    try:
         with connection.cursor() as cursor:
            cursor.execute(sql,params) 
            row = cursor.fetchone()
            return  row[0]
    except (Exception, psycopg2.DatabaseError) as error:
        raise error


list_existing_rows=[]

sql_existing_row="""
SELECT EXISTS(SELECT 1 FROM app_inventory
WHERE  serial_number<>'-' 
and serial_number=%(serial_param)s
and product_type_id=%(type_param)s 
and project_id = %(project_param)s 
)
"""
# SELECT EXISTS(SELECT 1 FROM app_inventory
# WHERE  (serial_number<>'-' and serial_number='FFGL2206A0FS-TEST'
# and  product_type_id=10 and  project_id = 23 ))

# init_param = {"serial_param":'FFGL2206A0FS-TEST' ,"type_param":10,"project_param":213}
# isExisting=is_existing_row(sql_existing_row,init_param,get_postgres_conn())
# print(isExisting)

for index,row in df_new.iterrows(): 
 init_param = {"serial_param":row['serial_number'] ,"type_param":row['product_type_id'],"project_param":row['project_id']}
 isExisting=is_existing_row(sql_existing_row,init_param,get_postgres_conn())
 if isExisting:
    list_existing_rows.append(row['temp_id'])
    


# In[176]:


df_existing=pd.DataFrame()
if len(list_existing_rows)>0:
 df_existing=df_new.query("temp_id in @list_existing_rows")
 df_new=df_new.drop(list_existing_rows, axis=0)
 df_existing= df_existing[main_cols_error_file]
 is_inccorect_data=True   


# In[177]:


df_existing


# In[178]:


#df_new


# In[179]:


#df_trasns.to_excel(tran2s_path2,index=False)
if is_inccorect_data:
  writer=pd.ExcelWriter(error_file,engine='xlsxwriter')
  if  not df_duplicatedRows.empty:
    df_duplicatedRows.to_excel(writer, sheet_name="Dupplicated_Rows",index=False)
  if  not pkNull_df.empty:
    pkNullCols_sr.to_excel(writer, sheet_name="NotFoundReferKey_Columns")
    pkNull_df.to_excel(writer, sheet_name="NotFoundReferKey_Rows",index=False)
  if  not df_existing.empty:
    df_existing.to_excel(writer, sheet_name="Existing_Rows",index=False)  
  writer.save()
else:
   print("No any incomplete data") 


# In[180]:


df_new=df_new.reset_index(drop=True)
df_new=df_new.drop(columns=['temp_id'])
df_new


# # Save to Database

# In[181]:


def nan_to_null_float(f,
        _NULL=psycopg2.extensions.AsIs('NULL'),
        _Float=psycopg2.extensions.Float):
    if not np.isnan(f):
        return _Float(f)
    return _NULL

psycopg2.extensions.register_adapter(float, nan_to_null_float)

def nan_to_null_int(f,
        _NULL=psycopg2.extensions.AsIs('NULL'),
        _Int=psycopg2.extensions.Int):
    if not np.isnan(f):
        return _Int(f)
    return _NULL

psycopg2.extensions.register_adapter(int, nan_to_null_int)

def add_data_values(df, table,conn):
    """
    Using psycopg2.extras.execute_values() to insert the dataframe
    """
    # Create a list of tupples from the dataframe values
    tuples = [tuple(x) for x in df.to_numpy()]
    # Comma-separated dataframe columns
    cols = ','.join(list(df.columns))
    # SQL quert to execute
    query  = "INSERT INTO %s(%s) VALUES %%s" % (table, cols)
    #print(query)
    #return query,tuples
    cursor = conn.cursor()
    try:
        extras.execute_values(cursor, query, tuples)
        conn.commit()
    except (Exception, psycopg2.DatabaseError) as error:
        print("Error: %s" % error)
        conn.rollback()
        cursor.close()
        raise error
        return 0
    
    return 1
    cursor.close()
    


# In[182]:


if df_new.empty==False: 
    result=add_data_values(df_new,'app_inventory',get_postgres_conn())
    if  result==1:
        print(f"{len(df_new.index)} items have been imported to database successfully.")
        print("importing data succeeded")
else:
    print("Error")
    print("No new inventory to import")
    print("All item exists in  inventory")


# In[183]:


if is_inccorect_data==False:
    os.remove(inventory_file)
    print(f"Import successfully so {os.path.abspath(inventory_file)} was deleted.")
    
else:
    print(f"Cannot import some inventories into database")
    print(f"check error file in {os.path.abspath(error_file)} compare to {os.path.abspath(inventory_file)}")


# In[184]:


#Test

# sql_cols_test="SELECT column_name FROM information_schema.columns WHERE  table_name   = 'test_inventoy'"
# listCols_TestTable= list_data(sql_cols_test ,None,get_postgres_conn())
# listCols_TestTable=listCols_TestTable['column_name'].tolist()
# listCols_TestTable.remove('id')
# print(listCols_TestTable)
# df_test=df_new[listCols_TestTable]
# df_test

# add_data_values(df_test,'test_inventoy',get_postgres_conn())


# In[ ]:





# In[ ]:





# In[ ]:




