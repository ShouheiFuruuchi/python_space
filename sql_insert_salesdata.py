from sql_parts import INSERT_SALESDATA

import pandas as pd
import datetime
from datetime import datetime
import numpy as np

tenpo = {
    "柏":"FUN柏",
    "千葉":"FUN千葉C-one",
    "伊勢崎":"FUNスマーク伊勢崎",
    "長町" :"FUNララガーデン長町",
    "船橋":"FUNららぽーとTOKYO-BAY",
    "富士見":"FUNららぽーと富士見",
    "レイク":"FUNイオンレイクタウン",
    "海老名":"FUNららぽーと海老名",
    "むさし":"FUNイオンモールむさし村山",
    "平塚":"FUNららぽーと湘南平塚",
    "名取":"FUNイオンモール名取",
    "大高":"FUNイオンモール大高",
    "東郷町":"FUNららぽーと愛知東郷",
    "太田":"FUNイオンモール太田",
    "水戸":"FUNイオンモール水戸内原",
    "エキスポ":"FUNららぽーとEXPOCITY",
    "川崎":"FUNラゾーナ川崎プラザ",
    "新三郷":"FUNららぽーと新三郷",
    "幕張":"FUNイオンモール幕張新都心",
    "各務原":"FUNイオンモール各務原",
    "堺" :"FUNららぽーと堺",
    
}

file_path = "C:/Users/古内翔平/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/シフト管理/予実管理.xlsx"

r_file = pd.read_excel(file_path,sheet_name="実績データ")
df_file = pd.DataFrame(r_file)

counter = 0
for row in df_file.values:
    
    # if counter == 300:
    #     break
    
    # else:
        
    counter_s = 0
    for shops in tenpo:
        try:
            
            if np.nan_to_num(row[1 + (counter_s * 7)]) > 0:
                print(str(row[0]))
                print(str(row[0])[0:4])
                print(str(row[0])[5:7])
                print(str(row[0])[8:10])
        
                print(shops)
                print(np.nan_to_num(row[1 + (counter_s * 7)]))
                print(np.nan_to_num(row[2+ (counter_s * 7)]))
                print(np.nan_to_num(row[3+ (counter_s * 7)]))
                print(np.nan_to_num(row[4+ (counter_s * 7)]))
                print(np.nan_to_num(row[5+ (counter_s * 7)]))
                print(np.nan_to_num(row[6+ (counter_s * 7)]))
                print(np.nan_to_num(row[7+ (counter_s * 7)]))
                    
            
            
            counter_s += 1
            
        except TypeError:
            print("NoType")    
            
    counter += 1        
    #INSERT_COMPETITOR(
    #