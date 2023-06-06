#競合店実績をDBへ登録

from sql_parts import INSERT_COMPETITOR

import pandas as pd

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

r_file = pd.read_excel(file_path,sheet_name="競合店実績")
df_file = pd.DataFrame(r_file)

for row in df_file.values:
    print(row)
    INSERT_COMPETITOR(
                      BRAND=str(row[0]),
                      SHOP_TYPE=tenpo[row[1]],
                      YEAR=int(row[2]),
                      MONTH=int(row[3]),
                      DAY=int(row[4]),
                      DOW=row[5],
                      DOW_TYPE=row[6],
                      SALES=row[7],
                      NOC=(row[8]),
                      )






