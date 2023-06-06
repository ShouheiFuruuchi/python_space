import openpyxl as xlpy
import pandas as pd
from selenium import webdriver
import time
from operator import itemgetter
import os
import shutil
import datetime
from datetime import timedelta,date
import time

#parts_1からインポート

from .parts_31 import priod_list #priod_list

#ーーーーーーーーーーーーーーーーーーーーー|　販売NETスクレイピング |ーーーーーーーーーーーーーーーーーーーーーーーーー

kasiwa = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[3]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[3]','柏','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[3]','柏.CSV']
chiba = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[4]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[4]', '千葉','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[4]','千葉.CSV']
isesaki = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[9]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[9]','伊勢崎','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[9]','伊勢崎.CSV']
nagamachi = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[11]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[11]','長町','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[11]','長町.CSV']
hunabashi = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[12]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[12]','船橋','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[12]','船橋.CSV']
hujimi = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[13]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[13]','富士見','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[13]','富士見.CSV']
reiku = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[15]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[15]','レイク','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[15]','レイク.CSV']
ebina = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[17]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[17]','海老名','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[17]','海老名.CSV']
musashi = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[18]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[18]','むさし','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[18]','むさし.CSV']
hiratuka = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[19]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[19]','平塚','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[19]','平塚.CSV']
natori = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[20]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[20]','名取','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[20]','名取.CSV']
otaka = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[21]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[21]','大高','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[21]','大高.CSV']
togocyo = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[22]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[22]','東郷町','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[22]','東郷町.CSV']
ota = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[23]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[23]','太田','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[23]','太田.CSV']
mito = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[24]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[24]','水戸','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[24]','水戸.CSV']
expo = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[25]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[25]','EXPO','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[25]','EXPO.CSV']
kawasaki = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[26]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[26]','川崎','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[26]','川崎.CSV']
sinmisato = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[27]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[27]','新三郷','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[27]','新三郷.CSV']
makuhari = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[28]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[28]','幕張','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[28]','幕張.CSV']
kagamihara = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[29]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[29]','各務原','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[29]','各務原.CSV']
sakai = ['//*[@id="ContentPlaceHolder1_DropDownListCond04"]/option[30]','//*[@id="ContentPlaceHolder1_DropDownListCond05"]/option[30]','堺','//*[@id="ContentPlaceHolder1_DropDownList9"]/option[30]','堺.CSV']

tenpo_list = [
  kasiwa,
  chiba,
  isesaki,
  # nagamachi,
  # hunabashi,
  hujimi,
  reiku,
  ebina,
  musashi,
  hiratuka,
  natori,
  otaka,
  togocyo,
  ota,
  mito,
  expo,
  kawasaki,
  sinmisato,
  makuhari,
  kagamihara,
  sakai,
  ]


tenpo = [
    ["01001008 FUN柏","柏"],
    ["01001009 FUN千葉C-one","千葉"],
    ["01001028 FUNスマーク伊勢崎","伊勢崎"],
    # ["01001032 FUNララガーデン長町","長町"],
    # ["01001033 FUNららぽーとTOKYO-BAY","船橋"],
    ["01001034 FUNららぽーと富士見","富士見"],
    ["01001036 FUNイオンレイクタウン","レイク"],
    ["01001038 FUNららぽーと海老名","海老名"],
    ["01001039 FUNイオンモールむさし村山","むさし"],
    ["01001040 FUNららぽーと湘南平塚","平塚"],
    ["01001041 FUNイオンモール名取","名取"],
    ["01001042 FUNイオンモール大高","大高"],
    ["01001043 FUNららぽーと愛知東郷","東郷町"],
    ["01001044 FUNイオンモール太田","太田"],
    ["01001045 FUNイオンモール水戸内原","水戸"],
    ["01001046 FUNららぽーとEXPOCITY","EXPO"],
    ["01001047 FUNラゾーナ川崎プラザ","川崎"],
    ["01001048 FUNららぽーと新三郷","新三郷"],
    ["01001049 FUNイオンモール幕張新都心","幕張"],
    ["01001050 FUNイオンモール各務原","各務原"],
    ["01001051 FUNららぽーと堺","堺"],
    
]

#-----------------------------------------------------------------------------------------------------------------------------
#　ここから　

dr_files = 'C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/data_folder'#今週実績
dr_files2 = 'C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/previous_data'#過去実績
dr_files3 = 'C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/sales_values'#売上集計

dr_read = os.listdir(dr_files)
dr_read2 = os.listdir(dr_files2)
dr_read3 = os.listdir(dr_files3)

print(dr_read)
print(dr_files2)
print(dr_files3)

#-----------------------------------------------------------------------------------------------------------------------------

#week1 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/basket-analysis/data-folder/品番売上集計データ.csv',encoding='SHIFT-JIS')#ok
#week2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/basket-analysis/data-folder/品番売上集計データ.csv',encoding='SHIFT-JIS')#ok
week1 = pd.read_csv('C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/data_folder/全店.csv',encoding='cp932')#今週実績
week1_sales_values = pd.read_csv('C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/sales_values/売上実績4.csv',encoding='cp932')#今週売上集計


previous_week1 = pd.read_csv("C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/previous_data/全店1.csv",encoding='cp932')#過去実績今週
previous_week1_sales_values = pd.read_csv('C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/sales_values/売上実績1.csv',encoding='cp932')#前週売上集計

previous_week2 = pd.read_csv("C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/previous_data/全店2.csv",encoding='cp932')#過去実績翌週
previous_week2_sales_values = pd.read_csv('C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/sales_values/売上実績2.csv',encoding='cp932')#今週売上集計

previous_week3 = pd.read_csv("C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/previous_data/全店3.csv",encoding='cp932')#過去実績翌週
previous_week3_sales_values = pd.read_csv('C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/sales_values/売上実績3.csv',encoding='cp932')#来週売上集計


output_faile = ["C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/週間分析.xlsx","週次","商品実績"]#パス/Sheet Name


df_week1 = pd.DataFrame(week1)#前週実績
df_week1_sales_values = pd.DataFrame(week1_sales_values)

print(df_week1)

noc = sum(df_week1_sales_values["売上客数"].values)#売上客数

item_cd = pd.DataFrame(df_week1["商品コード"].astype('str').str.zfill(10).values,columns=["商品CD"])
item_name = pd.DataFrame(df_week1["商品名"].values,columns=["商品名"])
category_cd = pd.DataFrame(df_week1["商品コード"].astype('str').str.zfill(10).str[2:4].values,columns=["アイテムCD"])
quantity = pd.DataFrame(df_week1['合計数量'].values,columns=["数量"])
amount = pd.DataFrame(df_week1['合計金額'].values,columns=["金額"])


df_week1_values = pd.concat([item_cd,item_name,category_cd,quantity,amount],axis=1)

filter1_df_week1_values = df_week1_values[df_week1_values["アイテムCD"] != "98" ]

filter2_df_week1_values = filter1_df_week1_values[(filter1_df_week1_values["商品名"] != "ｷﾚｲﾏｽｸ") & (filter1_df_week1_values["商品名"] != "ｻﾝﾌﾟﾙ") ]


all_amount = sum(filter2_df_week1_values["金額"].values)

op_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "01"]
op_1_amount = sum(op_1["金額"].values)

try :
  op_1_ratio = op_1_amount / all_amount
  
except ZeroDivisionError :
  op_1_ratio = 0

op_1_quant = sum(op_1["数量"].values)
op_1_demand = op_1_quant / noc#需要値

op_list = [op_1_amount,op_1_ratio,op_1_demand,op_1_quant]#

cd_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "02"]
cd_1_amount = sum(cd_1["金額"].values)

try :
  cd_1_ratio = cd_1_amount / all_amount 
  
except ZeroDivisionError :
  cd_1_ratio = 0

cd_1_quant = sum(cd_1["数量"].values)
cd_1_demand = cd_1_quant / noc#需要値

cd_list = [cd_1_amount,cd_1_ratio,cd_1_demand,cd_1_quant]#

jk_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "03"]
jk_1_amount = sum(jk_1["金額"].values)

try :
  jk_1_ratio = jk_1_amount / all_amount 
except ZeroDivisionError :
  jk_1_ratio = 0

jk_1_quant = sum(jk_1["数量"].values)
jk_1_demand = jk_1_quant / noc#需要値

jk_list = [jk_1_amount,jk_1_ratio,jk_1_demand,jk_1_quant]#

kt_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "04"]
kt_1_amount = sum(kt_1["金額"].values)

try :
  kt_1_ratio = kt_1_amount / all_amount 
except ZeroDivisionError :
  kt_1_ratio = 0

kt_1_quant = sum(kt_1["数量"].values)
kt_1_demand = kt_1_quant / noc#需要値
kt_list = [kt_1_amount,kt_1_ratio,kt_1_demand,kt_1_quant]#

cs_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "05"]
cs_1_amount = sum(cs_1["金額"].values)

try :
  cs_1_ratio = cs_1_amount / all_amount 
  
except ZeroDivisionError :
  cs_1_ratio = 0
  
cs_1_quant = sum(cs_1["数量"].values)
cs_1_demand = cs_1_quant / noc#需要値
cs_list = [cs_1_amount,cs_1_ratio,cs_1_demand,cs_1_quant]#

ct_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "06"]
ct_1_amount = sum(ct_1["金額"].values)

try :
  ct_1_ratio = ct_1_amount / all_amount 
except ZeroDivisionError :
  ct_1_ratio = 0

ct_1_quant = sum(ct_1["数量"].values)
ct_1_demand = ct_1_quant / noc#需要値
ct_list = [ct_1_amount,ct_1_ratio,ct_1_demand,ct_1_quant]#

bl_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "07"]
bl_1_amount = sum(bl_1["金額"].values)

try:
  bl_1_ratio = bl_1_amount / all_amount 
  
except ZeroDivisionError:
  
  bl_1_ratio = 0
  
bl_1_quant = sum(bl_1["数量"].values)
bl_1_demand = bl_1_quant / noc#需要値
bl_list = [bl_1_amount,bl_1_ratio,bl_1_demand,bl_1_quant]#

sk_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "08"]
sk_1_amount = sum(sk_1["金額"].values)

try: 
  sk_1_ratio = sk_1_amount / all_amount 
  
except ZeroDivisionError:
  
  sk_1_ratio = 0

sk_1_quant = sum(sk_1["数量"].values)
sk_1_demand = sk_1_quant / noc#需要値
sk_list = [sk_1_amount,sk_1_ratio,sk_1_demand,sk_1_quant]#

pt_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "09"]
pt_1_amount = sum(pt_1["金額"].values)

try :
  
  pt_1_ratio = pt_1_amount / all_amount 
  
except ZeroDivisionError:
  
  pt_1_ratio = 0
    

pt_1_quant = sum(pt_1["数量"].values)
pt_1_demand = pt_1_quant / noc#需要値
pt_list = [pt_1_amount,pt_1_ratio,pt_1_demand,pt_1_quant]#

tr_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "10"]
tr_1_amount = sum(tr_1["金額"].values)

try :
  
  tr_1_ratio = tr_1_amount / all_amount 
  
except ZeroDivisionError:
  
  tr_1_ratio = 0
  
tr_1_quant = sum(tr_1["数量"].values)
tr_1_demand = tr_1_quant / noc#需要値
tr_list = [tr_1_amount,tr_1_ratio,tr_1_demand,tr_1_quant]#

inn_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "11"]
inn_1_amount = sum(inn_1["金額"].values)

try :
  
  inn_1_ratio = inn_1_amount / all_amount 
  
except ZeroDivisionError:
  inn_1_ratio = 0
  
inn_1_quant = sum(inn_1["数量"].values)
inn_1_demand = inn_1_quant / noc#需要値
inn_list = [inn_1_amount,inn_1_ratio,inn_1_demand,inn_1_quant]#

setup_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "12"]
setup_1_amount = sum(setup_1["金額"].values)

try :
  setup_1_ratio = setup_1_amount / all_amount
  
except ZeroDivisionError:
  setup_1_ratio = 0

setup_1_quant = sum(setup_1["数量"].values)
setup_1_demand = setup_1_quant / noc#需要値
setup_list = [setup_1_amount,setup_1_ratio,setup_1_demand,setup_1_quant]#

acc_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "13"]
acc_1_amount = sum(acc_1["金額"].values)

try :
  acc_1_ratio = acc_1_amount / all_amount
  
except ZeroDivisionError:
  acc_1_ratio = 0

acc_1_quant = sum(acc_1["数量"].values)
acc_1_demand = acc_1_quant / noc#需要値
acc_list = [acc_1_amount,acc_1_ratio,acc_1_demand,acc_1_quant]#


sh_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "15"]
sh_1_amount = sum(sh_1["金額"].values)

try :
  sh_1_ratio = sh_1_amount / all_amount
  
except ZeroDivisionError:
  sh_1_ratio = 0

sh_1_quant = sum(sh_1["数量"].values)
sh_1_demand = sh_1_quant / noc#需要値
sh_list = [sh_1_amount,sh_1_ratio,sh_1_demand,sh_1_quant]#

out_put_list = [
  op_list,
  cd_list,
  jk_list,
  kt_list,
  cs_list,
  ct_list,
  bl_list,
  sk_list,
  pt_list,
  tr_list,
  inn_list,
  setup_list,
  acc_list,
  sh_list
]

#週間分析をファイル読み込み

out_wb = xlpy.load_workbook(output_faile[0])

#週間分析⇒週次を指定
out_ws = out_wb[output_faile[1]]

#----------------------------------------
header = 2 
header2 = 17

low = 0
low2 = 0
#----------------------------------------

for i in out_put_list:
  out_ws["C" + str(header + low)].value = i[0]
  out_ws["D" + str(header + low)].value = i[1]
  out_ws["E" + str(header + low)].value = i[2]
  
  low += 1
  
for i_name in tenpo:
  out_ws_2 = out_wb[i_name[1]]#シート名を指定
  
  #<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  #全店実績
  #<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  
  #OP/SETUP
  out_ws_2["K" + str(33)].value = out_put_list[0][0] + out_put_list[11][0]
  out_ws_2["L" + str(33)].value = out_put_list[0][1] + out_put_list[11][1]
  out_ws_2["M" + str(33)].value = out_put_list[0][2] + out_put_list[11][2]
  out_ws_2["J" + str(33)].value = out_put_list[0][3] + out_put_list[11][3]
  #TOPs
  out_ws_2["K" + str(34)].value = out_put_list[3][0] + out_put_list[4][0] + out_put_list[6][0] + out_put_list[9][0]
  out_ws_2["L" + str(34)].value = out_put_list[3][1] + out_put_list[4][1] + out_put_list[6][1] + out_put_list[9][1]
  out_ws_2["M" + str(34)].value = out_put_list[3][2] + out_put_list[4][2] + out_put_list[6][2] + out_put_list[9][2]
  out_ws_2["J" + str(34)].value = out_put_list[3][3] + out_put_list[4][3] + out_put_list[6][3] + out_put_list[9][3]
  
  #BOTTOMs
  out_ws_2["K" + str(35)].value = out_put_list[7][0] + out_put_list[8][0]
  out_ws_2["L" + str(35)].value = out_put_list[7][1] + out_put_list[8][1]
  out_ws_2["M" + str(35)].value = out_put_list[7][2] + out_put_list[8][2]
  out_ws_2["J" + str(35)].value = out_put_list[7][3] + out_put_list[8][3]
  
  #羽織
  out_ws_2["K" + str(36)].value = out_put_list[1][0] + out_put_list[2][0] + out_put_list[5][0]
  out_ws_2["L" + str(36)].value = out_put_list[1][1] + out_put_list[2][1] + out_put_list[5][1]
  out_ws_2["M" + str(36)].value = out_put_list[1][2] + out_put_list[2][2] + out_put_list[5][2]
  out_ws_2["J" + str(36)].value = out_put_list[1][3] + out_put_list[2][3] + out_put_list[5][3]
  
  #インナー
  out_ws_2["K" + str(37)].value = out_put_list[10][0]
  out_ws_2["L" + str(37)].value = out_put_list[10][1]
  out_ws_2["M" + str(37)].value = out_put_list[10][2]
  out_ws_2["J" + str(37)].value = out_put_list[10][3]
  
  #ACC
  out_ws_2["K" + str(38)].value = out_put_list[12][0]
  out_ws_2["L" + str(38)].value = out_put_list[12][1]
  out_ws_2["M" + str(38)].value = out_put_list[12][2]
  out_ws_2["J" + str(38)].value = out_put_list[12][3]
  
  #集計実績
  
  out_ws_2["K" + str(30)].value = out_put_list[0][0] + out_put_list[1][0] + out_put_list[2][0] + out_put_list[3][0] + out_put_list[4][0] + out_put_list[5][0] + out_put_list[6][0] + out_put_list[7][0] + out_put_list[8][0] + out_put_list[9][0] + out_put_list[10][0] + out_put_list[11][0] + out_put_list[12][0]
  
  
  out_ws_2["K" + str(39)].value = out_put_list[0][0] + out_put_list[1][0] + out_put_list[2][0] + out_put_list[3][0] + out_put_list[4][0] + out_put_list[5][0] + out_put_list[6][0] + out_put_list[7][0] + out_put_list[8][0] + out_put_list[9][0] + out_put_list[10][0] + out_put_list[11][0] + out_put_list[12][0]
  
  
  out_ws_2["L" + str(30)].value = out_put_list[0][1] + out_put_list[1][1] + out_put_list[2][1] + out_put_list[3][1] + out_put_list[4][1] + out_put_list[5][1] + out_put_list[6][1] + out_put_list[7][1] + out_put_list[8][1] + out_put_list[9][1] + out_put_list[10][1] + out_put_list[11][1] + out_put_list[12][1]

  
  out_ws_2["L" + str(39)].value = out_put_list[0][1] + out_put_list[1][1] + out_put_list[2][1] + out_put_list[3][1] + out_put_list[4][1] + out_put_list[5][1] + out_put_list[6][1] + out_put_list[7][1] + out_put_list[8][1] + out_put_list[9][1] + out_put_list[10][1] + out_put_list[11][1] + out_put_list[12][1]
  
  
  out_ws_2["M" + str(30)].value = out_put_list[0][2] + out_put_list[1][2] + out_put_list[2][2] + out_put_list[3][2] + out_put_list[4][2] + out_put_list[5][2] + out_put_list[6][2] + out_put_list[7][2] + out_put_list[8][2] + out_put_list[9][2] + out_put_list[10][2] + out_put_list[11][2] + out_put_list[12][2]
  
  out_ws_2["M" + str(39)].value = out_put_list[0][2] + out_put_list[1][2] + out_put_list[2][2] + out_put_list[3][2] + out_put_list[4][2] + out_put_list[5][2] + out_put_list[6][2] + out_put_list[7][2] + out_put_list[8][2] + out_put_list[9][2] + out_put_list[10][2] + out_put_list[11][2] + out_put_list[12][2]
  
  
  out_ws_2["J" + str(30)].value = out_put_list[0][3] + out_put_list[1][3] + out_put_list[2][3] + out_put_list[3][3] + out_put_list[4][3] + out_put_list[5][3] + out_put_list[6][3] + out_put_list[7][2] + out_put_list[8][3] + out_put_list[9][3] + out_put_list[10][3] + out_put_list[11][3] + out_put_list[12][3]
  
  out_ws_2["J" + str(39)].value = out_put_list[0][3] + out_put_list[1][3] + out_put_list[2][3] + out_put_list[3][3] + out_put_list[4][3] + out_put_list[5][3] + out_put_list[6][3] + out_put_list[7][3] + out_put_list[8][3] + out_put_list[9][3] + out_put_list[10][3] + out_put_list[11][3] + out_put_list[12][3]
  
  #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  #14 シューヅ込み
  for ii in range(0,13):
    ii_1 = out_put_list[ii]
    
    out_ws_2["K" + str(header2 + ii)].value = ii_1[0]
    out_ws_2["L" + str(header2 + ii)].value = ii_1[1]
    out_ws_2["M" + str(header2 + ii)].value = ii_1[2]
    out_ws_2["J" + str(header2 + ii)].value = ii_1[3]
    
   
      
    
out_wb.save(output_faile[0])

#前週実績出力完了
#★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
#前年実績を出力

#previous_week1 = pd.read_csv("C:/Users/fun-f/Desktop/analysis/previous_data/全店1.csv",encoding='SHIFT-JIS')#過去実績今週

previous_list =[
  previous_week1,
  previous_week2,
  previous_week3
]

previous_list_sales = [
  previous_week1_sales_values,
  previous_week2_sales_values,
  previous_week3_sales_values
]


column_list2 = [
  "C",
  "D",
  "E",
  "F",
  "G",
  "H",
  "I",
  "J",
  "K"
]

output_faile = ["C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/週間分析.xlsx","週次","商品実績"]#パス/Sheet Name
col_no = 0

for path_n,path_o in zip(previous_list,previous_list_sales) :


  df_previous_week1 = pd.DataFrame(path_n)
  df_previous_week1_values = pd.DataFrame(path_o)

  print(df_week1)
  
  previous_noc = sum(df_previous_week1_values["売上客数"].values)#売上客数

  previous_item_cd = pd.DataFrame(df_previous_week1["商品コード"].astype('str').str.zfill(10).values,columns=["商品CD"])
  previous_item_name = pd.DataFrame(df_previous_week1["商品名"].values,columns=["商品名"])
  previous_category_cd = pd.DataFrame(df_previous_week1["商品コード"].astype('str').str.zfill(10).str[2:4].values,columns=["アイテムCD"])
  previous_quantity = pd.DataFrame(df_previous_week1['合計数量'].values,columns=["数量"])
  previous_amount = pd.DataFrame(df_previous_week1['合計金額'].values,columns=["金額"])


  df_previous_week1_values = pd.concat([previous_item_cd,previous_item_name,previous_category_cd,previous_quantity,previous_amount],axis=1)

  filter1_df_previous_week1_values = df_previous_week1_values[df_previous_week1_values["アイテムCD"] != "98" ]

  filter2_df_previous_week1_values = filter1_df_previous_week1_values[(filter1_df_previous_week1_values["商品名"] != "ｷﾚｲﾏｽｸ") & (filter1_df_week1_values["商品名"] != "ｻﾝﾌﾟﾙ") ]


  previous_all_amount = sum(filter2_df_previous_week1_values["金額"].values)

  previous_op_1 = filter2_df_previous_week1_values[filter2_df_previous_week1_values["アイテムCD"] == "01"]
  previous_op_1_amount = sum(previous_op_1["金額"].values)
  try:
    previous_op_1_ratio = previous_op_1_amount / previous_all_amount 
    
  except ZeroDivisionError:
    previous_op_1_ratio = 0
      
  previous_op_1_quant = sum(previous_op_1["数量"].values)
  previous_op_1_demand = previous_op_1_quant / previous_noc#需要値
  previous_op_list = [previous_op_1_amount,previous_op_1_ratio,previous_op_1_demand]
  
  
  previous_cd_1 = filter2_df_previous_week1_values[filter2_df_previous_week1_values["アイテムCD"] == "02"]
  previous_cd_1_amount = sum(previous_cd_1["金額"].values)
  
  try:
    previous_cd_1_ratio = previous_cd_1_amount / previous_all_amount 
  except ZeroDivisionError:
    previous_cd_1_ratio = 0
    
  
  previous_cd_1_quant = sum(previous_cd_1["数量"].values)
  previous_cd_1_demand = previous_cd_1_quant / previous_noc#需要値
  previous_cd_list = [previous_cd_1_amount,previous_cd_1_ratio,previous_cd_1_demand]
  
  previous_jk_1 = filter2_df_previous_week1_values[filter2_df_previous_week1_values["アイテムCD"] == "03"]
  previous_jk_1_amount = sum(previous_jk_1["金額"].values)
  
  try :
    previous_jk_1_ratio = previous_jk_1_amount / previous_all_amount 
    
  except ZeroDivisionError:
    previous_jk_1_ratio = 0
  
  previous_jk_1_quant = sum(previous_jk_1["数量"].values)
  previous_jk_1_demand = previous_jk_1_quant / previous_noc#需要値
  previous_jk_list = [previous_jk_1_amount,previous_jk_1_ratio,previous_jk_1_demand]
  
  previous_kt_1 = filter2_df_previous_week1_values[filter2_df_previous_week1_values["アイテムCD"] == "04"]
  previous_kt_1_amount = sum(previous_kt_1["金額"].values)
  
  try:
    previous_kt_1_ratio = previous_kt_1_amount / previous_all_amount 
    
  except ZeroDivisionError:
    
    previous_kt_1_ratio = 0
  
  previous_kt_1_quant = sum(previous_kt_1["数量"].values)
  previous_kt_1_demand = previous_kt_1_quant / previous_noc#需要値
  previous_kt_list = [previous_kt_1_amount,previous_kt_1_ratio,previous_kt_1_demand]
  
  previous_cs_1 = filter2_df_previous_week1_values[filter2_df_previous_week1_values["アイテムCD"] == "05"]
  previous_cs_1_amount = sum(previous_cs_1["金額"].values)
  
  try:
    previous_cs_1_ratio = previous_cs_1_amount / previous_all_amount 
    
  except ZeroDivisionError:
    
    previous_cs_1_ratio = 0
    
  
  previous_cs_1_quant = sum(previous_cs_1["数量"].values)
  previous_cs_1_demand = previous_cs_1_quant / previous_noc#需要値
  previous_cs_list = [previous_cs_1_amount,previous_cs_1_ratio,previous_cs_1_demand]
  
  
  previous_ct_1 = filter2_df_previous_week1_values[filter2_df_previous_week1_values["アイテムCD"] == "06"]
  previous_ct_1_amount = sum(previous_ct_1["金額"].values)
  
  try:
    previous_ct_1_ratio = previous_ct_1_amount / previous_all_amount 
    
  except ZeroDivisionError:
    previous_ct_1_ratio = 0
    
  
  previous_ct_1_quant = sum(previous_ct_1["数量"].values)
  previous_ct_1_demand = previous_ct_1_quant / previous_noc#需要値
  previous_ct_list = [previous_ct_1_amount,previous_ct_1_ratio,previous_ct_1_demand]
  
  
  previous_bl_1 = filter2_df_previous_week1_values[filter2_df_previous_week1_values["アイテムCD"] == "07"]
  previous_bl_1_amount = sum(previous_bl_1["金額"].values)
  
  try :
    previous_bl_1_ratio = previous_bl_1_amount / previous_all_amount 
    
  except ZeroDivisionError:
    
    previous_bl_1_ratio = 0
    
  
  previous_bl_1_quant = sum(previous_bl_1["数量"].values)
  previous_bl_1_demand = previous_bl_1_quant / previous_noc#需要値
  previous_bl_list = [previous_bl_1_amount,previous_bl_1_ratio,previous_bl_1_demand]
  
  previous_sk_1 = filter2_df_previous_week1_values[filter2_df_previous_week1_values["アイテムCD"] == "08"]
  previous_sk_1_amount = sum(previous_sk_1["金額"].values)
  
  try:
    previous_sk_1_ratio = previous_sk_1_amount / previous_all_amount 
    
  except ZeroDivisionError:
    previous_sk_1_ratio = 0
    
  
  previous_sk_1_quant = sum(previous_sk_1["数量"].values)
  previous_sk_1_demand = previous_sk_1_quant / previous_noc#需要値
  previous_sk_list = [previous_sk_1_amount,previous_sk_1_ratio,previous_sk_1_demand]
  
  
  previous_pt_1 = filter2_df_previous_week1_values[filter2_df_previous_week1_values["アイテムCD"] == "09"]
  previous_pt_1_amount = sum(previous_pt_1["金額"].values)
  
  try :
    previous_pt_1_ratio = previous_pt_1_amount / previous_all_amount 
    
  except ZeroDivisionError:
    previous_pt_1_ratio = 0
  
  previous_pt_1_quant = sum(previous_pt_1["数量"].values)
  previous_pt_1_demand = previous_pt_1_quant / previous_noc#需要値
  previous_pt_list = [previous_pt_1_amount,previous_pt_1_ratio,previous_pt_1_demand]
  
  
  previous_tr_1 = filter2_df_previous_week1_values[filter2_df_previous_week1_values["アイテムCD"] == "10"]
  previous_tr_1_amount = sum(previous_tr_1["金額"].values)
  
  try:
    
    previous_tr_1_ratio = previous_tr_1_amount / previous_all_amount 
    
  except ZeroDivisionError:
    
    previous_tr_1_ratio = 0
    
  
  previous_tr_1_quant = sum(previous_tr_1["数量"].values)
  previous_tr_1_demand = previous_tr_1_quant / previous_noc#需要値
  previous_tr_list = [previous_tr_1_amount,previous_tr_1_ratio,previous_tr_1_demand]
  
  
  previous_inn_1 = filter2_df_previous_week1_values[filter2_df_previous_week1_values["アイテムCD"] == "11"]
  previous_inn_1_amount = sum(previous_inn_1["金額"].values)
  
  try:
    
    previous_inn_1_ratio = previous_inn_1_amount / previous_all_amount 
    
  except ZeroDivisionError:
    previous_inn_1_ratio = 0
    
  
  previous_inn_1_quant = sum(previous_inn_1["数量"].values)
  previous_inn_1_demand = previous_inn_1_quant / previous_noc#需要値
  previous_inn_list = [previous_inn_1_amount,previous_inn_1_ratio,previous_inn_1_demand]
  
  previous_setup_1 = filter2_df_previous_week1_values[filter2_df_previous_week1_values["アイテムCD"] == "12"]
  previous_setup_1_amount = sum(previous_setup_1["金額"].values)
  
  try:
    previous_setup_1_ratio = previous_setup_1_amount / previous_all_amount 
    
  except ZeroDivisionError:
    previous_setup_1_ratio = 0
  
  previous_setup_1_quant = sum(previous_setup_1["数量"].values)
  previous_setup_1_demand = previous_setup_1_quant / previous_noc#需要値
  previous_setup_list = [previous_setup_1_amount,previous_setup_1_ratio,previous_setup_1_demand]
  
  
  previous_acc_1 = filter2_df_previous_week1_values[filter2_df_previous_week1_values["アイテムCD"] == "13"]
  previous_acc_1_amount = sum(previous_acc_1["金額"].values)
  try:
    previous_acc_1_ratio = previous_acc_1_amount / previous_all_amount
    
  except ZeroDivisionError:
    previous_acc_1_ratio = 0
  
  previous_acc_1_quant = sum(previous_acc_1["数量"].values)
  previous_acc_1_demand = previous_acc_1_quant / previous_noc#需要値 
  previous_acc_list = [previous_acc_1_amount,previous_acc_1_ratio,previous_acc_1_demand]
  
  
  previous_sh_1 = filter2_df_previous_week1_values[filter2_df_previous_week1_values["アイテムCD"] == "15"]
  previous_sh_1_amount = sum(previous_sh_1["金額"].values)
  try:
    previous_sh_1_ratio = previous_sh_1_amount / previous_all_amount
    
  except ZeroDivisionError:
    previous_sh_1_ratio = 0
  
  previous_sh_1_quant = sum(previous_sh_1["数量"].values)
  previous_sh_1_demand = previous_sh_1_quant / previous_noc#需要値 
  previous_sh_list = [previous_sh_1_amount,previous_sh_1_ratio,previous_sh_1_demand]
  
  #-------------------------------------------------------------------------------------------------


  previous_out_put_list = [
    previous_op_list,#0
    previous_cd_list,#1
    previous_jk_list,#2
    previous_kt_list,#3
    previous_cs_list,#4
    previous_ct_list,#5
    previous_bl_list,#6
    previous_sk_list,#7
    previous_pt_list,#8
    previous_tr_list,#9
    previous_inn_list,#10
    previous_setup_list,
    previous_acc_list,
    previous_sh_list
  ]


  out_wb = xlpy.load_workbook(output_faile[0])

  out_ws = out_wb[output_faile[1]]

  #----------------------------------------
  header2 = 18
  low = 0
  #----------------------------------------
  #3/30日修正トリム平均
  #header2/3を修正＆追加
  
  #----------------------------------------
  #店別出力設定
  header3 = 44
  
  #----------------------------------------

  for i_1 in previous_out_put_list:
    out_ws[str(column_list2[0 + col_no]) + str(header2 + low)].value = i_1[0]
    out_ws[str(column_list2[1 + col_no]) + str(header2 + low)].value = i_1[1]
    out_ws[str(column_list2[2 + col_no]) + str(header2 + low)].value = i_1[2]
    
    out_ws[str(column_list2[1]) + str(16)].value = priod_list[0]
    out_ws[str(column_list2[4]) + str(16)].value = priod_list[1]
    out_ws[str(column_list2[7]) + str(16)].value = priod_list[2]
    
    out_ws[str(column_list2[3]) + str(2)].value = '=IFERROR(TRIMMEAN(各務原:柏!E' + str( 17 + col_no ) + ',週次!G'+ str(2 + col_no) + '-週次!H'+ str(2 + col_no) + '),0)'#トリム平均
    out_ws[str(column_list2[4]) + str(2)].value = '=IFERROR(AVERAGEA(各務原:柏!E' + str( 17 + col_no ) + '),0)'#平均
    out_ws[str(column_list2[5]) + str(2)].value = '=IFERROR(STDEV.P(各務原:柏!E' + str( 17 + col_no ) + '),0)'#標準偏差
    
    
    for i_10 in tenpo:
      out_ws_shop = out_wb[i_10[1]]
      
      #アイテム別実績集計
      
      out_ws_shop[str(column_list2[0 + col_no]) + str(header3 + low)].value = i_1[0]
      out_ws_shop[str(column_list2[1 + col_no]) + str(header3 + low)].value = i_1[1]
      out_ws_shop[str(column_list2[2 + col_no]) + str(header3 + low)].value = i_1[2]
      
      #用途区分実績集計
    
      index_1 = 60 #OP/SETUP
    
      out_ws_shop[str(column_list2[0 + col_no]) + str(index_1)].value = previous_out_put_list[0][0] + previous_out_put_list[11][0]#売上
      out_ws_shop[str(column_list2[1 + col_no]) + str(index_1)].value = previous_out_put_list[0][1] + previous_out_put_list[11][1]#構成比【売上】
      out_ws_shop[str(column_list2[2 + col_no]) + str(index_1)].value = previous_out_put_list[0][2] + previous_out_put_list[11][2]#購入率
      
      
      index_2 = 61 #TOPs
    
      out_ws_shop[str(column_list2[0 + col_no]) + str(index_2)].value = previous_out_put_list[3][0] + previous_out_put_list[4][0] + previous_out_put_list[6][0] + previous_out_put_list[9][0]#売上
      out_ws_shop[str(column_list2[1 + col_no]) + str(index_2)].value = previous_out_put_list[3][1] + previous_out_put_list[4][1] + previous_out_put_list[6][1] + previous_out_put_list[9][1]#構成比【売上】
      out_ws_shop[str(column_list2[2 + col_no]) + str(index_2)].value = previous_out_put_list[3][2] + previous_out_put_list[4][2] + previous_out_put_list[6][2] + previous_out_put_list[9][2]#購入率
      
      
      index_3 = 62 #BOTTOMs
    
      out_ws_shop[str(column_list2[0 + col_no]) + str(index_3)].value = previous_out_put_list[7][0] + previous_out_put_list[8][0]#売上
      out_ws_shop[str(column_list2[1 + col_no]) + str(index_3)].value = previous_out_put_list[7][1] + previous_out_put_list[8][1]#構成比【売上】
      out_ws_shop[str(column_list2[2 + col_no]) + str(index_3)].value = previous_out_put_list[7][2] + previous_out_put_list[8][2]#購入率
      
      index_4 = 63 #羽織
    
      out_ws_shop[str(column_list2[0 + col_no]) + str(index_4)].value = previous_out_put_list[1][0] + previous_out_put_list[2][0] + previous_out_put_list[5][0]#売上
      out_ws_shop[str(column_list2[1 + col_no]) + str(index_4)].value = previous_out_put_list[1][1] + previous_out_put_list[2][1] + previous_out_put_list[5][1]#構成比【売上】
      out_ws_shop[str(column_list2[2 + col_no]) + str(index_4)].value = previous_out_put_list[1][2] + previous_out_put_list[2][2] + previous_out_put_list[5][2]#購入率
      
      index_5 = 64 #INNER
    
      out_ws_shop[str(column_list2[0 + col_no]) + str(index_5)].value = previous_out_put_list[10][0]#売上
      out_ws_shop[str(column_list2[1 + col_no]) + str(index_5)].value = previous_out_put_list[10][1]#構成比【売上】
      out_ws_shop[str(column_list2[2 + col_no]) + str(index_5)].value = previous_out_put_list[10][2]#購入率
      
      index_6 = 65 #ACC
    
      out_ws_shop[str(column_list2[0 + col_no]) + str(index_6)].value = previous_out_put_list[12][0]#売上
      out_ws_shop[str(column_list2[1 + col_no]) + str(index_6)].value = previous_out_put_list[12][1]#構成比【売上】
      out_ws_shop[str(column_list2[2 + col_no]) + str(index_6)].value = previous_out_put_list[12][2]#購入率
      
      index_7 = 66 #集計
      
      
      out_ws_shop[str(column_list2[0 + col_no]) + str(57)].value = previous_out_put_list[0][0] + previous_out_put_list[1][0] + previous_out_put_list[2][0] + previous_out_put_list[3][0] + previous_out_put_list[4][0] + previous_out_put_list[5][0] + previous_out_put_list[6][0] + previous_out_put_list[7][0] + previous_out_put_list[8][0] + previous_out_put_list[9][0] + previous_out_put_list[10][0] + previous_out_put_list[11][0] + previous_out_put_list[12][0]#売上
    
      out_ws_shop[str(column_list2[0 + col_no]) + str(index_7)].value = previous_out_put_list[0][0] + previous_out_put_list[1][0] + previous_out_put_list[2][0] + previous_out_put_list[3][0] + previous_out_put_list[4][0] + previous_out_put_list[5][0] + previous_out_put_list[6][0] + previous_out_put_list[7][0] + previous_out_put_list[8][0] + previous_out_put_list[9][0] + previous_out_put_list[10][0] + previous_out_put_list[11][0] + previous_out_put_list[12][0]#売上
      
      
      out_ws_shop[str(column_list2[1 + col_no]) + str(57)].value = previous_out_put_list[0][1] + previous_out_put_list[1][1] + previous_out_put_list[2][1] + previous_out_put_list[3][1] + previous_out_put_list[4][1] + previous_out_put_list[5][1] + previous_out_put_list[6][1] + previous_out_put_list[7][1] + previous_out_put_list[8][1] + previous_out_put_list[9][1] + previous_out_put_list[10][1] + previous_out_put_list[11][1] + previous_out_put_list[12][1]#構成比【売上】
      
      out_ws_shop[str(column_list2[1 + col_no]) + str(index_7)].value = previous_out_put_list[0][1] + previous_out_put_list[1][1] + previous_out_put_list[2][1] + previous_out_put_list[3][1] + previous_out_put_list[4][1] + previous_out_put_list[5][1] + previous_out_put_list[6][1] + previous_out_put_list[7][1] + previous_out_put_list[8][1] + previous_out_put_list[9][1] + previous_out_put_list[10][1] + previous_out_put_list[11][1] + previous_out_put_list[12][1]#構成比【売上】
      
      
      out_ws_shop[str(column_list2[2 + col_no]) + str(57)].value = previous_out_put_list[0][2] + previous_out_put_list[1][2] + previous_out_put_list[2][2] + previous_out_put_list[3][2] + previous_out_put_list[4][2] + previous_out_put_list[5][2] + previous_out_put_list[6][1] + previous_out_put_list[7][2] + previous_out_put_list[8][2] + previous_out_put_list[9][2] + previous_out_put_list[10][2] + previous_out_put_list[11][2] + previous_out_put_list[12][2]#購入率
      
      out_ws_shop[str(column_list2[2 + col_no]) + str(index_7)].value = previous_out_put_list[0][2] + previous_out_put_list[1][2] + previous_out_put_list[2][2] + previous_out_put_list[3][2] + previous_out_put_list[4][2] + previous_out_put_list[5][2] + previous_out_put_list[6][1] + previous_out_put_list[7][2] + previous_out_put_list[8][2] + previous_out_put_list[9][2] + previous_out_put_list[10][2] + previous_out_put_list[11][2] + previous_out_put_list[12][2]#購入率
      
      
      
      out_ws[str(column_list2[1]) + str(42)].value = priod_list[0]
      out_ws[str(column_list2[4]) + str(42)].value = priod_list[1]
      out_ws[str(column_list2[7]) + str(42)].value = priod_list[2]
      
      out_ws_shop[str(column_list2[1]) + str(42)].value = priod_list[0]
      out_ws_shop[str(column_list2[4]) + str(42)].value = priod_list[1]
      out_ws_shop[str(column_list2[7]) + str(42)].value = priod_list[2]
      
    
    low += 1


  col_no += 3
  out_wb.save(output_faile[0])
  
  
  


#★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
#----------------------------------------
header = 2 
low = 0
#----------------------------------------

out_ws2 = out_wb[output_faile[2]]

sort_file2 = filter2_df_week1_values.sort_values("数量",ascending=False)
sort_file1 = filter2_df_week1_values.sort_values("金額",ascending=False)

#-----------------------------------------
#ここを一度解放★★★
#週間分析ファイルに出力
for j,k,l,m,n in zip(sort_file1["商品CD"].values,sort_file1["商品名"].values,sort_file1["アイテムCD"].values,sort_file1["数量"].values,sort_file1["金額"].values):
  
  if low == 70 :#ここを修正★
    print("Non")

  else:
    out_ws2["A" + str(header + low)].value = j
    out_ws2["B" + str(header + low)].value = k
    out_ws2["C" + str(header + low)].value = l
    out_ws2["D" + str(header + low)].value = m
    out_ws2["E" + str(header + low)].value = n
    
    low += 1

out_wb.save(output_faile[0])

data_concat_list = []

#def analysis1():
  
for select_i in tenpo_list:
  
  shop_i = select_i
    
    
  week2 = pd.read_csv('C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/data_folder/' + str(shop_i[4]),encoding='cp932')#ok
  
  df_week1 = pd.DataFrame(week2)

  print(df_week1)

  item_cd = pd.DataFrame(df_week1["商品コード"].astype('str').str.zfill(10).values,columns=["商品CD"])
  item_name = pd.DataFrame(df_week1["商品名"].values,columns=["商品名"])
  category_cd = pd.DataFrame(df_week1["商品コード"].astype('str').str.zfill(10).str[2:4].values,columns=["アイテムCD"])
  quantity = pd.DataFrame(df_week1['合計数量'].values,columns=["数量"])
  amount = pd.DataFrame(df_week1['合計金額'].values,columns=["金額"])
  #shop_name = pd.DataFrame([shop_i[2]],columns=["店舗"])

  df_week1_values = pd.concat([item_cd,item_name,category_cd,quantity,amount],axis=1)

  filter1_df_week1_values = df_week1_values[df_week1_values["アイテムCD"] != "98" ]

  filter2_df_week1_values = filter1_df_week1_values[(filter1_df_week1_values["商品名"] != "ｷﾚｲﾏｽｸ") & (filter1_df_week1_values["商品名"] != "ｻﾝﾌﾟﾙ") ]
  
  filter2_df_week1_values2 = pd.DataFrame(filter2_df_week1_values)
  
  sort_filter2_df_week1_values2 = filter2_df_week1_values2.sort_values("数量",ascending=False)#★★★ 修正 ★★★
  #sort_filter2_df_week1_values3 = filter2_df_week1_values2.sort_values("金額",ascending=False)#★★★ 修正 ★★★
  
  for low_data in sort_filter2_df_week1_values2.values :#★★★ 修正 ★★★
    item_cd = pd.DataFrame([low_data[0]],columns=["商品CD"])
    item_name = pd.DataFrame([low_data[1]],columns=["商品名"])
    category_cd = pd.DataFrame([low_data[2]],columns=["アイテムCD"])
    quantity = pd.DataFrame([low_data[3]],columns=["数量"])
    amount = pd.DataFrame([low_data[4]],columns=["金額"])
    shop_name = pd.DataFrame([shop_i[2]],columns=["店舗"])
    
    low_data2 = pd.concat([item_cd,item_name,category_cd,quantity,amount,shop_name],axis=1)
    
    data_concat_list.append(low_data2)
    

  all_amount = sum(filter2_df_week1_values["金額"].values)

  op_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "01"]
  op_1_amount = sum(op_1["金額"].values)
  try :
    op_1_ratio = op_1_amount / all_amount 
    
  except ZeroDivisionError:
    
    op_1_ratio = 0
    
  op_list = [op_1_amount,op_1_ratio]

  cd_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "02"]
  cd_1_amount = sum(cd_1["金額"].values)
  
  try:
    cd_1_ratio = cd_1_amount / all_amount 
    
  except ZeroDivisionError:
    cd_1_ratio = 0
    
  cd_list = [cd_1_amount,cd_1_ratio]
  
  
  jk_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "03"]
  jk_1_amount = sum(jk_1["金額"].values)
  try:
    jk_1_ratio = jk_1_amount / all_amount 
    
  except ZeroDivisionError:
    jk_1_ratio = 0
    
  jk_list = [jk_1_amount,jk_1_ratio]

  kt_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "04"]
  kt_1_amount = sum(kt_1["金額"].values)
  
  try :
    kt_1_ratio = kt_1_amount / all_amount 
    
  except ZeroDivisionError:
    kt_1_ratio = 0
    
  kt_list = [kt_1_amount,kt_1_ratio]

  cs_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "05"]
  cs_1_amount = sum(cs_1["金額"].values)
  
  try :
    cs_1_ratio = cs_1_amount / all_amount 
    
  except ZeroDivisionError:
    cs_1_ratio = 0
    
  cs_list = [cs_1_amount,cs_1_ratio]

  ct_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "06"]
  ct_1_amount = sum(ct_1["金額"].values)
  
  try:
    ct_1_ratio = ct_1_amount / all_amount 
    
  except ZeroDivisionError:
    ct_1_ratio = 0
    
  ct_list = [ct_1_amount,ct_1_ratio]

  bl_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "07"]
  bl_1_amount = sum(bl_1["金額"].values)
  
  try :
    bl_1_ratio = bl_1_amount / all_amount 
    
  except ZeroDivisionError :
    bl_1_ratio = 0
    
  bl_list = [bl_1_amount,bl_1_ratio]

  sk_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "08"]
  sk_1_amount = sum(sk_1["金額"].values)
  
  try :
    sk_1_ratio = sk_1_amount / all_amount 
    
  except ZeroDivisionError:
    sk_1_ratio = 0
    
  sk_list = [sk_1_amount,sk_1_ratio]

  pt_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "09"]
  pt_1_amount = sum(pt_1["金額"].values)
  
  try :
    pt_1_ratio = pt_1_amount / all_amount 
    
  except ZeroDivisionError:
    pt_1_ratio = 0
    
  pt_list = [pt_1_amount,pt_1_ratio]

  tr_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "10"]
  tr_1_amount = sum(tr_1["金額"].values)
  
  try :
    tr_1_ratio = tr_1_amount / all_amount 
    
  except ZeroDivisionError:
    
    tr_1_ratio = 0
    
  tr_list = [tr_1_amount,tr_1_ratio]

  inn_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "11"]
  inn_1_amount = sum(inn_1["金額"].values)
  
  try:
    inn_1_ratio = inn_1_amount / all_amount 
    
  except ZeroDivisionError:
    inn_1_ratio = 0
    
  inn_list = [inn_1_amount,inn_1_ratio]

  setup_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "12"]
  setup_1_amount = sum(setup_1["金額"].values)
  
  try:
    setup_1_ratio = setup_1_amount / all_amount
    
  except ZeroDivisionError:
    setup_1_ratio = 0
    
  setup_list = [setup_1_amount,setup_1_ratio]

  acc_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "13"]
  acc_1_amount = sum(acc_1["金額"].values)
  
  try:
    acc_1_ratio = acc_1_amount / all_amount
    
  except ZeroDivisionError:
    
    acc_1_ratio = 0
    
  acc_list = [acc_1_amount,acc_1_ratio]
  
  
  sh_1 = filter2_df_week1_values[filter2_df_week1_values["アイテムCD"] == "15"]
  sh_1_amount = sum(sh_1["金額"].values)
  
  try:
    sh_1_ratio = sh_1_amount / all_amount
    
  except ZeroDivisionError:
    
    sh_1_ratio = 0
    
  sh_list = [sh_1_amount,sh_1_ratio]

  out_put_list = [
    
    op_list,
    cd_list,
    jk_list,
    kt_list,
    cs_list,
    ct_list,
    bl_list,
    sk_list,
    pt_list,
    tr_list,
    inn_list,
    setup_list,
    acc_list,
    sh_list
    
  ]
  
  print(all_amount)


  
print(data_concat_list)  

#data_concat_list2 = pd.DataFrame([data_concat_list])

try :

  df_data_concat_list = pd.concat(data_concat_list,axis=0)
  
except ValueError:  
  
  item_cd = pd.DataFrame([""],columns=["商品CD"])
  item_name = pd.DataFrame([""],columns=["商品名"])
  category_cd = pd.DataFrame([""],columns=["アイテムCD"])
  quantity = pd.DataFrame([0],columns=["数量"])
  amount = pd.DataFrame([0],columns=["金額"])
  shop_name = pd.DataFrame([""],columns=["店舗"])
  
  low_data2 = pd.concat([item_cd,item_name,category_cd,quantity,amount,shop_name],axis=1)
  
  data_concat_list.append(low_data2)
  
  
  
  df_data_concat_list = pd.concat(data_concat_list,axis=0)
  
  
  
print(df_data_concat_list)


column_list = {
  1:["F","G"],2:["H","I"],3:["J","K"],4:["L","M"],5:["N","O"],6:["P","Q"],7:["R","S"],8:["T","U"],9:["V","W"],10:["X","Y"],11:["Z","AA"],12:["AB","AC"],13:["AD","AE"],14:["AF","AG"],15:["AH","AI"],16:["AJ","AK"],17:["AL","AM"],18:["AN","AO"],19:["AP","AQ"],20:["AR","AS"]
}
#★★★ ここに移動【3/11】 ★★★
#-----------------------------------------
#週間分析ファイルに出力
low = 0
data_count = 0
rank_count = 1
index_no = 2
for j,k,l,m,n in zip(sort_file1["商品CD"].values,sort_file1["商品名"].values,sort_file1["アイテムCD"].values,sort_file1["数量"].values,sort_file1["金額"].values):
  if low == 70 :#ここを修正★3.15
    break
  
  else:
    out_ws2["A" + str(header + low)].value = j
    out_ws2["B" + str(header + low)].value = k
    out_ws2["C" + str(header + low)].value = l
    out_ws2["D" + str(header + low)].value = m
    out_ws2["E" + str(header + low)].value = n
    
    rank_data = df_data_concat_list[df_data_concat_list["商品CD"] == j]
    sort_rank_data = rank_data.sort_values("数量",ascending=False)
    #sort_rank_data = rank_data.sort_values("金額",ascending=False)
    print(sort_rank_data)
    
    #for outdata in sort_rank_data :
    #shop_ = sort_rank_data["店舗"].values[data_count]
    #quantity_ = sort_rank_data["数量"].values[data_count]
    
    rank_count1 = 1
    try:
    
      shop_ = sort_rank_data["店舗"].values[rank_count1-1]
      quantity_ = sort_rank_data["数量"].values[rank_count1-1]
      
      out_ws2[ str(column_list[rank_count1][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count1][1]) + str(index_no) ].value = quantity_
    
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count1][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count1][1]) + str(index_no) ].value = ""
    
    
    rank_count2 = 2
    try:
    
      shop_ = sort_rank_data["店舗"].values[rank_count2-1]
      quantity_ = sort_rank_data["数量"].values[rank_count2-1]
      
      out_ws2[ str(column_list[rank_count2][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count2][1]) + str(index_no) ].value = quantity_
    
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count2][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count2][1]) + str(index_no) ].value = ""
    
    rank_count3 = 3
    try:
    
      shop_ = sort_rank_data["店舗"].values[rank_count3-1]
      quantity_ = sort_rank_data["数量"].values[rank_count3-1]
      
      out_ws2[ str(column_list[rank_count3][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count3][1]) + str(index_no) ].value = quantity_
    
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count3][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count3][1]) + str(index_no) ].value = ""
    
    rank_count4 = 4
    
    try:
    
      shop_ = sort_rank_data["店舗"].values[rank_count4-1]
      quantity_ = sort_rank_data["数量"].values[rank_count4-1]
      out_ws2[ str(column_list[rank_count4][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count4][1]) + str(index_no) ].value = quantity_
    
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count4][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count4][1]) + str(index_no) ].value = ""
    
    rank_count5 = 5
    try:
      shop_ = sort_rank_data["店舗"].values[rank_count5-1]
      quantity_ = sort_rank_data["数量"].values[rank_count5-1]
      out_ws2[ str(column_list[rank_count5][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count5][1]) + str(index_no) ].value = quantity_
    
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count5][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count5][1]) + str(index_no) ].value = ""
    
    rank_count6 = 6
    try:
      shop_ = sort_rank_data["店舗"].values[rank_count6-1]
      quantity_ = sort_rank_data["数量"].values[rank_count6-1]
      out_ws2[ str(column_list[rank_count6][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count6][1]) + str(index_no) ].value = quantity_
    
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count6][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count6][1]) + str(index_no) ].value = ""
    
    rank_count7 = 7
    try:
      shop_ = sort_rank_data["店舗"].values[rank_count7-1]
      quantity_ = sort_rank_data["数量"].values[rank_count7-1]
      out_ws2[ str(column_list[rank_count7][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count7][1]) + str(index_no) ].value = quantity_
    
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count7][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count7][1]) + str(index_no) ].value = ""
    
    rank_count8 = 8
    try:
      shop_ = sort_rank_data["店舗"].values[rank_count8-1]
      quantity_ = sort_rank_data["数量"].values[rank_count8-1]
      out_ws2[ str(column_list[rank_count8][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count8][1]) + str(index_no) ].value = quantity_
    
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count8][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count8][1]) + str(index_no) ].value = ""
    
    rank_count9 = 9
    try:
      shop_ = sort_rank_data["店舗"].values[rank_count9-1]
      quantity_ = sort_rank_data["数量"].values[rank_count9-1]
      out_ws2[ str(column_list[rank_count9][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count9][1]) + str(index_no) ].value = quantity_
    
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count9][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count9][1]) + str(index_no) ].value = ""
    
    rank_count10 = 10
    try:
      shop_ = sort_rank_data["店舗"].values[rank_count10-1]
      quantity_ = sort_rank_data["数量"].values[rank_count10-1]
      out_ws2[ str(column_list[rank_count10][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count10][1]) + str(index_no) ].value = quantity_
    
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count10][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count10][1]) + str(index_no) ].value = ""
    
    rank_count11 = 11
    try:
      shop_ = sort_rank_data["店舗"].values[rank_count11-1]
      quantity_ = sort_rank_data["数量"].values[rank_count11-1]
      out_ws2[ str(column_list[rank_count11][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count11][1]) + str(index_no) ].value = quantity_
    
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count11][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count11][1]) + str(index_no) ].value = ""
    
    rank_count12 = 12
    try:
      shop_ = sort_rank_data["店舗"].values[rank_count12-1]
      quantity_ = sort_rank_data["数量"].values[rank_count12-1]
      out_ws2[ str(column_list[rank_count12][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count12][1]) + str(index_no) ].value = quantity_
    
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count12][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count12][1]) + str(index_no) ].value = ""
    
    rank_count13 = 13
    try:
      shop_ = sort_rank_data["店舗"].values[rank_count13-1]
      quantity_ = sort_rank_data["数量"].values[rank_count13-1]
      out_ws2[ str(column_list[rank_count13][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count13][1]) + str(index_no) ].value = quantity_
    
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count13][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count13][1]) + str(index_no) ].value = ""
    
    rank_count14 = 14
    try:
      shop_ = sort_rank_data["店舗"].values[rank_count14-1]
      quantity_ = sort_rank_data["数量"].values[rank_count14-1]
      out_ws2[ str(column_list[rank_count14][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count14][1]) + str(index_no) ].value = quantity_
    
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count14][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count14][1]) + str(index_no) ].value = ""
      
    
    rank_count15= 15
    try:
      shop_ = sort_rank_data["店舗"].values[rank_count15-1]
      quantity_ = sort_rank_data["数量"].values[rank_count15-1]
      out_ws2[ str(column_list[rank_count15][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count15][1]) + str(index_no) ].value = quantity_
      
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count15][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count15][1]) + str(index_no) ].value = ""
    
    rank_count16 = 16
    try:
      shop_ = sort_rank_data["店舗"].values[rank_count16-1]
      quantity_ = sort_rank_data["数量"].values[rank_count16-1]
      out_ws2[ str(column_list[rank_count16][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count16][1]) + str(index_no) ].value = quantity_
    
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count16][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count16][1]) + str(index_no) ].value = ""
    
    rank_count17 = 17
    try:
      shop_ = sort_rank_data["店舗"].values[rank_count17-1]
      quantity_ = sort_rank_data["数量"].values[rank_count17-1]
      out_ws2[ str(column_list[rank_count17][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count17][1]) + str(index_no) ].value = quantity_
      
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count17][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count17][1]) + str(index_no) ].value = ""
    
    rank_count18 = 18
    try:
      shop_ = sort_rank_data["店舗"].values[rank_count18-1]
      quantity_ = sort_rank_data["数量"].values[rank_count18-1]
      out_ws2[ str(column_list[rank_count18][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count18][1]) + str(index_no) ].value = quantity_
      
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count18][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count18][1]) + str(index_no) ].value = ""
      
    
    rank_count19 = 19
    try :
      
      shop_ = sort_rank_data["店舗"].values[rank_count19-1]
      quantity_ = sort_rank_data["数量"].values[rank_count19-1]
      out_ws2[ str(column_list[rank_count19][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count19][1]) + str(index_no) ].value = quantity_
      
    except IndexError:
      
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count19][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count19][1]) + str(index_no) ].value = ""
      
    
    rank_count20 = 20
    
    try :
      
      shop_ = sort_rank_data["店舗"].values[rank_count20-1]
      quantity_ = sort_rank_data["数量"].values[rank_count20-1]
      out_ws2[ str(column_list[rank_count20][0]) + str(index_no) ].value = shop_
      out_ws2[ str(column_list[rank_count20][1]) + str(index_no) ].value = quantity_
      
    except IndexError :
      shop_ = ""
      quantity_ = ""
      out_ws2[ str(column_list[rank_count20][0]) + str(index_no) ].value = ""
      out_ws2[ str(column_list[rank_count20][1]) + str(index_no) ].value = ""

    
    
    data_count += 1
    rank_count += 1
    index_no += 1
      
    
    low += 1
    time.sleep(2)
    
    

    out_wb.save('C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/週間分析.xlsx')
      #'C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/週間分析.xlsx')  
    
    
    #--------------------------------------------------------------------------------------------------------
 