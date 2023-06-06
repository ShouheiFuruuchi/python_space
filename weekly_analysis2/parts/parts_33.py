import pandas as pd
import openpyxl as xlpy
#from parts_32 import column_list2

import os

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
  nagamachi,
  hunabashi,
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

print(tenpo[0][0][9:])

output_faile = ["C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/週間分析.xlsx","週次","商品実績","Sheet2"]#パス/Sheet Name
#-----------------------------------------------------------------------------------------------------------------------------
#　ここから　


week1_sales_values = pd.read_csv('C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/sales_values/売上実績4.csv',encoding='cp932')#今週売上集計

dr_files = 'C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/data_folder'#今週実績

dr_files2 = 'C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/previous_data'#過去実績
dr_files3 = 'C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/sales_values'#売上集計

#5月29日追記
#店舗在庫状況を出力
inventory_folder = "C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/inventory"
#》＞＞＞＞＞＞＞＞ 全店在庫の出力 - START - ＜＜＜＜＜＜＜＜＜＜＜＜＜

r_file = pd.read_csv(os.path.join(inventory_folder,'全店.csv'),encoding="cp932")

df_r_file = pd.DataFrame(r_file)


item_cd = pd.DataFrame(df_r_file["商品コード"].astype(str).str.zfill(10).values,columns=["商品CD"])  
item_name = pd.DataFrame(df_r_file["商品名"].values,columns=["商品名"])
category_cd = pd.DataFrame(df_r_file["商品コード"].astype(str).str.zfill(10).str[2:4].values,columns=["アイテムCD"])

inventory_quantity = pd.DataFrame(df_r_file["現在数量"].values,columns=["在庫数量"])
inventory_value = pd.DataFrame(df_r_file["原価合計"].values,columns=["在庫金額"])
color = pd.DataFrame(df_r_file["色名"].values,columns=["色名"])
size = pd.DataFrame(df_r_file["サイズ名"].values,columns=["サイズ名"])

inventory_list = pd.concat([item_cd,item_name,category_cd,color,size,inventory_quantity,inventory_value],axis=1)
print(inventory_list)

item_cd_list = [
  "01",#OP
  "02",#CD
  "03",#JK
  "04",#KT
  "05",#CS
  "06",#CT
  "07",#BL
  "08",#SK
  "09",#PT
  "10",#TR
  "11",#INN
  "12",#SETUP
  "13",#ACC
]

item_category = {
  "01":"OP",
  "02":"CD",
  "03":"JK",
  "04":"KT",
  "05":"CS",
  "06":"CT",
  "07":"BL",
  "08":"SK",
  "09":"PT",
  "10":"TR",
  "11":"INN",
  "12":"SETUP",
  "13":"ACC",
}

all_quantity = sum(inventory_list["在庫数量"].values)
print(all_quantity)
all_value = sum(inventory_list["在庫金額"].values)

inventory_list_2 = []
for i_cd in item_cd_list:
  print(item_category[i_cd])

  item_key = inventory_list[inventory_list["アイテムCD"] == i_cd ]
  print(item_key)
  
  #在庫点数を出力
  inventory_item_cd = pd.DataFrame([i_cd],columns=["アイテムCD"])
  inventory_item_category = pd.DataFrame([item_category[i_cd]],columns=["アイテムCD"])
  item_quantity = pd.DataFrame([sum(item_key["在庫数量"].values)],columns=["在庫数量"])
  print(item_quantity)
  item_quantity_ratio = pd.DataFrame([float(item_quantity.values/all_quantity)],columns=["在庫構成比 (数量)"])
  
  #在庫金額出力
  item_value = pd.DataFrame([sum(item_key["在庫金額"].values)],columns=["在庫金額"])
  item_value_ratio = pd.DataFrame([float(item_value.values/all_value)],columns=["在庫構成比 (金額)"])
  
  
  inventory_index = pd.concat([inventory_item_cd,inventory_item_category,item_quantity,item_quantity_ratio,item_value,item_value_ratio],axis=1)
  inventory_list_2.append(inventory_index)
  
inventory_list_concat = pd.concat(inventory_list_2,axis=0)

for shop_name_n in tenpo:
  out_wb = xlpy.load_workbook(output_faile[0])

  out_ws = out_wb[shop_name_n[1]]

  #----------------------------------------
  header = 17
  low = 0
  #----------------------------------------
  

  for inv_i in inventory_list_concat.values:

    out_ws["N" + str(header + low)].value = inv_i[2]
    out_ws["O" + str(header + low)].value = inv_i[3]
    
    low += 1
    
  #集計実績
    
  out_ws["N" + str(30)].value = sum(inventory_list_concat["在庫数量"])
  
  out_ws["O" + str(30)].value = sum(inventory_list_concat["在庫構成比 (数量)"])
  
  #<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  #全店実績
  #<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  
  #OP/SETUP
  out_ws["N" + str(33)].value = inventory_list_concat.values[0][2] + inventory_list_concat.values[11][2]
  out_ws["O" + str(33)].value = inventory_list_concat.values[0][3] + inventory_list_concat.values[11][3]
  
  
  #TOPs

  out_ws["N" + str(34)].value = inventory_list_concat.values[3][2] + inventory_list_concat.values[4][2] + inventory_list_concat.values[6][2] + inventory_list_concat.values[9][2]
  out_ws["O" + str(34)].value = inventory_list_concat.values[3][3] + inventory_list_concat.values[4][3] + inventory_list_concat.values[6][3] + inventory_list_concat.values[9][3]
  
  #BOTTOMs

  out_ws["N" + str(35)].value = inventory_list_concat.values[7][2] + inventory_list_concat.values[8][2]
  out_ws["O" + str(35)].value = inventory_list_concat.values[7][3] + inventory_list_concat.values[8][3]
  
  #羽織
  out_ws["N" + str(36)].value = inventory_list_concat.values[1][2] + inventory_list_concat.values[2][2] + inventory_list_concat.values[5][2]
  out_ws["O" + str(36)].value = inventory_list_concat.values[1][3] + inventory_list_concat.values[2][3] + inventory_list_concat.values[5][3]
  
  #インナー

  out_ws["N" + str(37)].value = inventory_list_concat.values[10][2]
  out_ws["O" + str(37)].value = inventory_list_concat.values[10][3]
  
  #ACC
 
  out_ws["N" + str(38)].value = inventory_list_concat.values[12][2]
  out_ws["O" + str(38)].value = inventory_list_concat.values[12][3]
  
  
  #各データ合計値
  data_no1 = 0
  data_no2 = 1
  data_no3 = 2
  
  #売上実績欄に記入
  
  
  #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  
  

  out_wb.save("C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/週間分析.xlsx")


#》＞＞＞＞＞＞＞＞ 全店在庫の出力　- END -　 ＜＜＜＜＜＜＜＜＜＜＜＜＜


for shop_name_n in tenpo:

  shop_ana1 = pd.read_csv(dr_files + '/' + str(shop_name_n[1]) + '.csv',encoding='cp932')


  df_week1 = pd.DataFrame(shop_ana1)#前週実績
  df_week1_sales_values = pd.DataFrame(week1_sales_values)

  print(df_week1)

  shop_select_name = df_week1_sales_values[df_week1_sales_values["拠点名"] == str(shop_name_n[0][9:])]#売上客数

  noc = shop_select_name["売上客数"].values[0]
  print(noc)
  buget_n = shop_select_name["売上予算"].values[0]

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
    acc_list
  ]
  
  #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  #5/29追記　ここから
  #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  

  r_file = pd.read_csv(os.path.join(inventory_folder,str(shop_name_n[1]) + '.csv'),encoding="cp932")

  df_r_file = pd.DataFrame(r_file)
  
  
  item_cd = pd.DataFrame(df_r_file["商品コード"].astype(str).str.zfill(10).values,columns=["商品CD"])  
  item_name = pd.DataFrame(df_r_file["商品名"].values,columns=["商品名"])
  category_cd = pd.DataFrame(df_r_file["商品コード"].astype(str).str.zfill(10).str[2:4].values,columns=["アイテムCD"])
  
  inventory_quantity = pd.DataFrame(df_r_file["現在数量"].values,columns=["在庫数量"])
  inventory_value = pd.DataFrame(df_r_file["原価合計"].values,columns=["在庫金額"])
  color = pd.DataFrame(df_r_file["色名"].values,columns=["色名"])
  size = pd.DataFrame(df_r_file["サイズ名"].values,columns=["サイズ名"])
  
  inventory_list = pd.concat([item_cd,item_name,category_cd,color,size,inventory_quantity,inventory_value],axis=1)
  print(inventory_list)
  
  item_cd_list = [
    "01",#OP
    "02",#CD
    "03",#JK
    "04",#KT
    "05",#CS
    "06",#CT
    "07",#BL
    "08",#SK
    "09",#PT
    "10",#TR
    "11",#INN
    "12",#SETUP
    "13",#ACC
    "15",#SH
  ]
  
  item_category = {
    "01":"OP",
    "02":"CD",
    "03":"JK",
    "04":"KT",
    "05":"CS",
    "06":"CT",
    "07":"BL",
    "08":"SK",
    "09":"PT",
    "10":"TR",
    "11":"INN",
    "12":"SETUP",
    "13":"ACC",
    "15":"SH",
  }
  
  all_quantity = sum(inventory_list["在庫数量"].values)
  print(all_quantity)
  all_value = sum(inventory_list["在庫金額"].values)
  
  inventory_list_2 = []
  for i_cd in item_cd_list:
    print(item_category[i_cd])
  
    item_key = inventory_list[inventory_list["アイテムCD"] == i_cd ]
    print(item_key)
    
    #在庫点数を出力
    inventory_item_cd = pd.DataFrame([i_cd],columns=["アイテムCD"])
    inventory_item_category = pd.DataFrame([item_category[i_cd]],columns=["アイテムCD"])
    item_quantity = pd.DataFrame([sum(item_key["在庫数量"].values)],columns=["在庫数量"])
    print(item_quantity)
    item_quantity_ratio = pd.DataFrame([float(item_quantity.values/all_quantity)],columns=["在庫構成比 (数量)"])
    
    #在庫金額出力
    item_value = pd.DataFrame([sum(item_key["在庫金額"].values)],columns=["在庫金額"])
    item_value_ratio = pd.DataFrame([float(item_value.values/all_value)],columns=["在庫構成比 (金額)"])
    
    
    inventory_index = pd.concat([inventory_item_cd,inventory_item_category,item_quantity,item_quantity_ratio,item_value,item_value_ratio],axis=1)
    inventory_list_2.append(inventory_index)
    
  inventory_list_concat = pd.concat(inventory_list_2,axis=0)
    
    
  #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  #5/29追記　ここから
  #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  
  out_wb = xlpy.load_workbook(output_faile[0])

  out_ws = out_wb[shop_name_n[1]]

  #----------------------------------------
  header = 17
  low = 0
  #----------------------------------------
  

  for i,inv_i in zip(out_put_list,inventory_list_concat.values):
    print(inv_i)
    out_ws["D" + str(1)].value = shop_name_n[0][9:]
    out_ws["D" + str(10)].value = buget_n
    out_ws["D" + str(12)].value = "=IFERROR(C3/C2,0)"
    out_ws["D" + str(header + low)].value = i[0]
    out_ws["E" + str(header + low)].value = i[1]
    out_ws["F" + str(header + low)].value = i[2]
    out_ws["C" + str(header + low)].value = i[3]
    out_ws["G" + str(header + low)].value = inv_i[2]
    out_ws["H" + str(header + low)].value = inv_i[3]
    
    low += 1
    
  #集計実績
  out_ws["D" + str(30)].value = out_put_list[0][0] + out_put_list[1][0] + out_put_list[2][0] + out_put_list[3][0] + out_put_list[4][0] + out_put_list[5][0] + out_put_list[6][0] + out_put_list[7][0] + out_put_list[8][0] + out_put_list[9][0] + out_put_list[10][0] + out_put_list[11][0] + out_put_list[12][0]
  out_ws["E" + str(30)].value = out_put_list[0][1] + out_put_list[1][1] + out_put_list[2][1] + out_put_list[3][1] + out_put_list[4][1] + out_put_list[5][1] + out_put_list[6][1] + out_put_list[7][1] + out_put_list[8][1] + out_put_list[9][1] + out_put_list[10][1] + out_put_list[11][1] + out_put_list[12][1]
  out_ws["F" + str(30)].value = out_put_list[0][2] + out_put_list[1][2] + out_put_list[2][2] + out_put_list[3][2] + out_put_list[4][2] + out_put_list[5][2] + out_put_list[6][2] + out_put_list[7][2] + out_put_list[8][2] + out_put_list[9][2] + out_put_list[10][2] + out_put_list[11][2] + out_put_list[12][2]
  out_ws["C" + str(30)].value = out_put_list[0][3] + out_put_list[1][3] + out_put_list[2][3] + out_put_list[3][3] + out_put_list[4][3] + out_put_list[5][3] + out_put_list[6][3] + out_put_list[7][3] + out_put_list[8][3] + out_put_list[9][3] + out_put_list[10][3] + out_put_list[11][3] + out_put_list[12][3]
    
  out_ws["G" + str(30)].value = sum(inventory_list_concat["在庫数量"])
  
  out_ws["H" + str(30)].value = sum(inventory_list_concat["在庫構成比 (数量)"])
  
  #<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  #全店実績
  #<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  
  #OP/SETUP
  out_ws["D" + str(33)].value = out_put_list[0][0] + out_put_list[11][0]
  out_ws["E" + str(33)].value = out_put_list[0][1] + out_put_list[11][1]
  out_ws["F" + str(33)].value = out_put_list[0][2] + out_put_list[11][2]
  out_ws["C" + str(33)].value = out_put_list[0][3] + out_put_list[11][3]
  out_ws["G" + str(33)].value = inventory_list_concat.values[0][2] + inventory_list_concat.values[11][2]
  out_ws["H" + str(33)].value = inventory_list_concat.values[0][3] + inventory_list_concat.values[11][3]
  
  
  #TOPs
  out_ws["D" + str(34)].value = out_put_list[3][0] + out_put_list[4][0] + out_put_list[6][0] + out_put_list[9][0]
  out_ws["E" + str(34)].value = out_put_list[3][1] + out_put_list[4][1] + out_put_list[6][1] + out_put_list[9][1]
  out_ws["F" + str(34)].value = out_put_list[3][2] + out_put_list[4][2] + out_put_list[6][2] + out_put_list[9][2]
  out_ws["C" + str(34)].value = out_put_list[3][3] + out_put_list[4][3] + out_put_list[6][3] + out_put_list[9][3]
  out_ws["G" + str(34)].value = inventory_list_concat.values[3][2] + inventory_list_concat.values[4][2] + inventory_list_concat.values[6][2] + inventory_list_concat.values[9][2]
  out_ws["H" + str(34)].value = inventory_list_concat.values[3][3] + inventory_list_concat.values[4][3] + inventory_list_concat.values[6][3] + inventory_list_concat.values[9][3]
  
  #BOTTOMs
  out_ws["D" + str(35)].value = out_put_list[7][0] + out_put_list[8][0]
  out_ws["E" + str(35)].value = out_put_list[7][1] + out_put_list[8][1]
  out_ws["F" + str(35)].value = out_put_list[7][2] + out_put_list[8][2]
  out_ws["C" + str(35)].value = out_put_list[7][3] + out_put_list[8][3]
  out_ws["G" + str(35)].value = inventory_list_concat.values[7][2] + inventory_list_concat.values[8][2]
  out_ws["H" + str(35)].value = inventory_list_concat.values[7][3] + inventory_list_concat.values[8][3]
  
  #羽織
  out_ws["D" + str(36)].value = out_put_list[1][0] + out_put_list[2][0] + out_put_list[5][0]
  out_ws["E" + str(36)].value = out_put_list[1][1] + out_put_list[2][1] + out_put_list[5][1]
  out_ws["F" + str(36)].value = out_put_list[1][2] + out_put_list[2][2] + out_put_list[5][2]
  out_ws["C" + str(36)].value = out_put_list[1][3] + out_put_list[2][3] + out_put_list[5][3]
  out_ws["G" + str(36)].value = inventory_list_concat.values[1][2] + inventory_list_concat.values[2][2] + inventory_list_concat.values[5][2]
  out_ws["H" + str(36)].value = inventory_list_concat.values[1][3] + inventory_list_concat.values[2][3] + inventory_list_concat.values[5][3]
  
  #インナー
  out_ws["D" + str(37)].value = out_put_list[10][0]
  out_ws["E" + str(37)].value = out_put_list[10][1]
  out_ws["F" + str(37)].value = out_put_list[10][2]
  out_ws["C" + str(37)].value = out_put_list[10][3]
  out_ws["G" + str(37)].value = inventory_list_concat.values[10][2]
  out_ws["H" + str(37)].value = inventory_list_concat.values[10][3]
  
  #ACC
  out_ws["D" + str(38)].value = out_put_list[12][0]
  out_ws["E" + str(38)].value = out_put_list[12][1]
  out_ws["F" + str(38)].value = out_put_list[12][2]
  out_ws["C" + str(38)].value = out_put_list[12][3]
  out_ws["G" + str(38)].value = inventory_list_concat.values[12][2]
  out_ws["H" + str(38)].value = inventory_list_concat.values[12][3]
  
  
  #各データ合計値
  data_no1 = 0
  data_no2 = 1
  data_no3 = 2
  
  #売上実績欄に記入
  out_ws["D" + str(11)].value = out_put_list[0][data_no1] + out_put_list[1][data_no1] + out_put_list[2][data_no1] + out_put_list[3][data_no1] + out_put_list[4][data_no1] +  out_put_list[5][data_no1] + out_put_list[6][data_no1] + out_put_list[7][data_no1] + out_put_list[8][data_no1] + out_put_list[9][data_no1] + out_put_list[10][data_no1] + out_put_list[11][data_no1] + out_put_list[12][data_no1]
  
  bug_1 = int(out_ws["D" + str(10)].value)
  
  #達成率
  try:
    out_ws["D" + str(12)].value = int(out_put_list[0][data_no1] + out_put_list[1][data_no1] + out_put_list[2][data_no1] + out_put_list[3][data_no1] + out_put_list[4][data_no1] +  out_put_list[5][data_no1] + out_put_list[6][data_no1] + out_put_list[7][data_no1] + out_put_list[8][data_no1] + out_put_list[9][data_no1] + out_put_list[10][data_no1] + out_put_list[11][data_no1] + out_put_list[12][data_no1]) / bug_1
  except ZeroDivisionError:
    
    out_ws["D" + str(12)].value = 0
      
    
  out_ws["D" + str(39)].value = out_put_list[0][data_no1] + out_put_list[1][data_no1] + out_put_list[2][data_no1] + out_put_list[3][data_no1] + out_put_list[4][data_no1] +  out_put_list[5][data_no1] + out_put_list[6][data_no1] + out_put_list[7][data_no1] + out_put_list[8][data_no1] + out_put_list[9][data_no1] + out_put_list[10][data_no1] + out_put_list[11][data_no1] + out_put_list[12][data_no1]
  
  out_ws["E" + str(39)].value = out_put_list[0][data_no2] + out_put_list[1][data_no2] + out_put_list[2][data_no2] + out_put_list[3][data_no2] + out_put_list[4][data_no2] + out_put_list[5][data_no2] + out_put_list[6][data_no2] + out_put_list[7][data_no2] + out_put_list[8][data_no2] + out_put_list[9][data_no2] + out_put_list[10][data_no2] + out_put_list[11][data_no2] + out_put_list[12][data_no2]
  
  out_ws["F" + str(39)].value = out_put_list[0][data_no3] + out_put_list[1][data_no3] + out_put_list[2][data_no3] + out_put_list[3][data_no3] + out_put_list[4][data_no3] + out_put_list[5][data_no3] + out_put_list[6][data_no3] + out_put_list[7][data_no3] + out_put_list[8][data_no3] + out_put_list[9][data_no3] + out_put_list[10][data_no3] + out_put_list[11][data_no3] + out_put_list[12][data_no3]
  
  
  #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  
  

  out_wb.save('C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/週間分析.xlsx')




