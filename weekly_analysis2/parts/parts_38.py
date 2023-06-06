import openpyxl as pyxl
import os



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

col_list = [
["B","C"],
["D","E"],
["F","G"],
["H","I"],
["J","K"],
["L","M"],
["N","O"],
["P","Q"],
["R","S"],
["T","U"],
["V","W"],
["X","Y"],
["Z","AA"],
["AB","AC"],
["AD","AE"],
["AF","AG"],
["AH","AI"],
["AJ","AK"],
["AL","AM"],
["AN","AO"],
["AP","AQ"],

]
path = "C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/4⃣販売部/古内/analysis/create_file/"

target_file = os.listdir(path)

r_file_path = os.path.join(path,target_file[0])

wb = pyxl.load_workbook(r_file_path)

#出力先ファイル
output_ws = wb["改善カテゴリー集計"]

for sheet_name,select_col in zip(tenpo,col_list) :
  shop_ws = wb[sheet_name[1]]
  
  Sales_Ratio = shop_ws["D12"].value #達成率
  
  output_ws[select_col[1] + str(2)].value = Sales_Ratio

  
  col = "R"
  
  element_1 = shop_ws[str(col) + "5"].value
  element_2 = shop_ws[str(col) + "6"].value
  element_3 = shop_ws[str(col) + "7"].value
  element_4 = shop_ws[str(col) + "8"].value
  element_5 = shop_ws[str(col) + "9"].value
  
  element_list =[element_1,element_2,element_3,element_4,element_5]
  
  for i,element in zip(range(5,10),element_list):
    output_ws[select_col[0] + str(i)].value = element
  
  
  
  col = "S"
  
  element_1 = shop_ws[str(col) + "5"].value
  element_2 = shop_ws[str(col) + "6"].value
  element_3 = shop_ws[str(col) + "7"].value
  element_4 = shop_ws[str(col) + "8"].value
  element_5 = shop_ws[str(col) + "9"].value
  
  element_list =[element_1,element_2,element_3,element_4,element_5]
  
  for i,element in zip(range(5,10),element_list):
    output_ws[select_col[1] + str(i)].value = element
  
  
  col = "E"
  
  element_1 = shop_ws[str(col) + "17"].value
  element_2 = shop_ws[str(col) + "18"].value
  element_3 = shop_ws[str(col) + "19"].value
  element_4 = shop_ws[str(col) + "20"].value
  element_5 = shop_ws[str(col) + "21"].value
  element_6 = shop_ws[str(col) + "22"].value
  element_7 = shop_ws[str(col) + "23"].value
  element_8 = shop_ws[str(col) + "24"].value
  element_9 = shop_ws[str(col) + "25"].value
  element_10 = shop_ws[str(col) + "26"].value
  element_11 = shop_ws[str(col) + "27"].value
  element_12 = shop_ws[str(col) + "28"].value
  element_13 = shop_ws[str(col) + "29"].value
  
  element_list =[element_1,element_2,element_3,element_4,element_5,element_6,element_7,element_8,element_9,element_10,element_11,element_12,element_13]
  
  for i,element in zip(range(12,25),element_list):
    output_ws[select_col[0] + str(i)].value = element
  
  col = "F"
  
  element_1 = shop_ws[str(col) + "17"].value
  element_2 = shop_ws[str(col) + "18"].value
  element_3 = shop_ws[str(col) + "19"].value
  element_4 = shop_ws[str(col) + "20"].value
  element_5 = shop_ws[str(col) + "21"].value
  element_6 = shop_ws[str(col) + "22"].value
  element_7 = shop_ws[str(col) + "23"].value
  element_8 = shop_ws[str(col) + "24"].value
  element_9 = shop_ws[str(col) + "25"].value
  element_10 = shop_ws[str(col) + "26"].value
  element_11 = shop_ws[str(col) + "27"].value
  element_12 = shop_ws[str(col) + "28"].value
  element_13 = shop_ws[str(col) + "29"].value
  
  element_list =[element_1,element_2,element_3,element_4,element_5,element_6,element_7,element_8,element_9,element_10,element_11,element_12,element_13]
  
  for i,element in zip(range(12,25),element_list):
    output_ws[select_col[1] + str(i)].value = element
  
  

wb.save(r_file_path)